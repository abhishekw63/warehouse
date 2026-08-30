web: gunicorn renee_cosmetics.wsgi --bind 0.0.0.0:$PORT --workers 1 --threads 8 --timeout 300
release: python manage.py migrate --noinput
