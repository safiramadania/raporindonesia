from pathlib import Path
import json
import cloudpickle
import numpy as np
import pandas as pd
import torch


def load_predictor(bundle_dir, device=None):
    bundle_dir = Path(bundle_dir)
    manifest = json.loads((bundle_dir / 'manifest.json').read_text(encoding='utf-8'))
    device = device or ('cuda' if torch.cuda.is_available() else 'cpu')
    with (bundle_dir / 'preprocessor.cloudpickle').open('rb') as f:
        preprocessor = cloudpickle.load(f)
    members = []
    for seed in manifest['ensemble_seeds']:
        ckpt = torch.load(bundle_dir / f'ft_ple64_seed_{seed}.pt', map_location=device, weights_only=False)
        members.append((ckpt['model'].to(device).eval(), np.asarray(ckpt['y_mean']), np.asarray(ckpt['y_scale'])))

    def predict(frame):
        required = [manifest['school_id_column'], *manifest['feature_columns']]
        missing = [column for column in required if column not in frame.columns]
        if missing:
            raise ValueError(f'Kolom input belum lengkap: {missing}')
        schools = frame[manifest['school_id_column']]
        prepared = preprocessor.transform(frame[manifest['feature_columns']].copy(), schools)
        x_num = torch.as_tensor(prepared['num'], dtype=torch.float32, device=device)
        x_cat = torch.as_tensor(prepared['cat'], dtype=torch.long, device=device)
        predictions = []
        with torch.no_grad():
            for model, y_mean, y_scale in members:
                scaled = model(x_num, x_cat).cpu().numpy()
                predictions.append(scaled * y_scale + y_mean)
        return pd.DataFrame(np.mean(predictions, axis=0), columns=manifest['targets'], index=frame.index)

    return predict, manifest
