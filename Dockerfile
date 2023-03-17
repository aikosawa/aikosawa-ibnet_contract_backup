# mojimojiのbuildがgccに依存するため、slim-busterではなくbusterを選択

FROM python:3.9.6-buster 

WORKDIR /app

COPY requirements.txt .
RUN python -m venv /venv && . /venv/bin/activate && pip install -r requirements.txt

COPY app ./app
COPY templates ./templates
COPY main.py .

CMD . /venv/bin/activate && python main.py
