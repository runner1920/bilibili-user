FROM python:3.9

ADD . /code
WORKDIR /code

ENV LANG zh_CN.UTF-8
ENV LC_ALL zh_CN.UTF-8
ENV TZ=Asia/Shanghai

# 下载相关依赖
RUN pip install -i https://pypi.tuna.tsinghua.edu.cn/simple  --default-timeout=60 --no-cache-dir -r docker/requirements.txt
CMD [ "python", "/code/xiaofang.py" ]


