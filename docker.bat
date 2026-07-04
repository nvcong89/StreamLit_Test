FROM python:3.9

WORKD** /app

COPY . .

RUN**ip install -r requirements**xt

EXPOSE**501

CMD ["**reamlit", "**n", "app.py",**--server.port=8501** "--server.address=0**.0.0"]
``**