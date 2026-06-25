@echo off
cd /d "%~dp0"
echo Updating dashboard...
python rebuild_tracker.py
echo Running data audit...
python audit_tracker.py
echo Pushing to GitHub...
git add sales_order_tracker.html
git commit -m "Auto-update dashboard"
git push
echo.
echo Done! Dashboard updated at:
echo https://guitar174.github.io/vv-dashbaord-frame/sales_order_tracker.html
pause
