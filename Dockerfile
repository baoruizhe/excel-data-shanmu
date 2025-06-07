# 使用Wine在Linux中模拟Windows环境
FROM ubuntu:20.04

# 避免交互式安装
ENV DEBIAN_FRONTEND=noninteractive

# 安装必要的包
RUN apt-get update && apt-get install -y \
    software-properties-common \
    wget \
    gnupg2

# 添加Wine仓库
RUN wget -nc https://dl.winehq.org/wine-builds/winehq.key && \
    apt-key add winehq.key && \
    add-apt-repository 'deb https://dl.winehq.org/wine-builds/ubuntu/ focal main'

# 安装Wine和Python
RUN apt-get update && apt-get install -y \
    winehq-stable \
    python3 \
    python3-pip \
    xvfb

# 设置工作目录
WORKDIR /app

# 复制Python文件
COPY excel.py .
COPY requirements.txt .

# 安装Python依赖
RUN pip3 install -r requirements.txt
RUN pip3 install pyinstaller

# 初始化Wine
RUN xvfb-run -a winecfg

# 构建脚本
COPY build.sh .
RUN chmod +x build.sh

CMD ["./build.sh"] 