FROM python:3

EXPOSE 5000

WORKDIR /APP

COPY . .

RUN pip install -r requiriments.txt

RUN pip install gunicorn

CMD ["gunicorn","-w","4","-b","0.0.0.0:5000","app:app"]
