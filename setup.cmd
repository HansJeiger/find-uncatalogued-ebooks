if not exist .env (
    copy .env.dist .env
)
call npm install
pause
exit