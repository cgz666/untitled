# 使用官方 Python 3.8 镜像作为基础镜像
FROM python:3.8-slim

# 设置工作目录为 /app
WORKDIR /app

# 将当前目录下的所有文件复制到容器的工作目录
COPY . /app

# 设置 PyPI 镜像源为阿里云
ENV PIP_INDEX_URL=https://mirrors.aliyun.com/pypi/simple/

# 安装依赖
RUN pip install --no-cache-dir -r requirements.txt

# 暴露 5000 端口
EXPOSE 5000

# 启动应用
CMD ["python", "app.py"]