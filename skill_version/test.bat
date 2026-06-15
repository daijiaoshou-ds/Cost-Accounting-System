@echo off
pushd %~dp0scripts
echo === Step 1: Core Calculation ===
python pipeline_cli.py --config ../test_config.json --output-dir ../output
echo.
echo === Step 2: Cost Fluctuation ===
python cost_fluctuation.py --output-dir ../output
echo.
echo === Step 3: Margin Analysis ===
python margin_analysis.py --output-dir ../output
echo.
echo === Step 4: Generate HTML Report ===
python generate_report.py --output-dir ../output
echo.
echo === Done ===
start "" "..\output\cost_report.html"
popd
