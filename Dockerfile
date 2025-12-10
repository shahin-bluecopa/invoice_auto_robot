FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

COPY libs/bluecopa_rpa_python-0.1.0-py3-none-any.whl /app/bluecopa_rpa_python-0.1.0-py3-none-any.whl
RUN pip install /app/*.whl

COPY . .
ENV ROBOT_ENTRYPOINT="python main.py"
CMD ["python", "main.py"]
