FROM python:3.11

WORKDIR /app
COPY . .

RUN python -m pip install --upgrade pip
RUN pip install -r requirements.txt

EXPOSE 3000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "3000"]
