@echo off
pushd %~dp0scripts
echo === Phase 0: Environment Check ===
python check_env.py --json
echo.
echo === Phase 1-5: Run All (skip env, already checked) ===
python run_all.py --config ../test_config.json --output-dir ../output --skip-env
echo.
echo === Done. Opening report... ===
start "" "..\output\cost_report.html"
popd
