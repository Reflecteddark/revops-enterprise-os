FROM python:3.11-slim

WORKDIR /app

COPY . /app

ENV PORT=10000
ENV PYTHONUNBUFFERED=1

EXPOSE 10000

CMD ["python", "services/max_bot.py"]
