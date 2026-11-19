FT-Transformer + PLE-64 deployment bundle

The ensemble contains three models (seeds 42, 17, 2025) and predicts LIT and NUM jointly.

Input DataFrame must include kd_sekolah and every column listed in manifest.json -> feature_columns. Do not pass kd_siswa_an or target columns as predictors. Missing predictor values are handled by the fitted school-aware preprocessor; unseen schools use its global fallback statistics.

Python used to train: 3.13.15. Install requirements.txt in a compatible environment.

Example:
from inference import load_predictor
predict, manifest = load_predictor('ft_ple64_deployment_bundle')
predictions = predict(input_df)

The saved test metrics and model parameters are in manifest.json.
