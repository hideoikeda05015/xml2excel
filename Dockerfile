FROM python:3.9@sha256:f3ea0f7b3662c33b650f763a365b2c3e3c7695ab1658fd0987150708b1f5f0e6

COPY requirements.txt .

RUN pip3 install --upgrade pip && \
    pip3 install -r requirements.txt

# https://qiita.com/SHinGo-Koba/items/a73e5f345c22c5ebbf96
RUN mkdir -p /usr/share/man/man1
RUN apt-get update 
RUN apt-get install -y git default-jre graphviz fonts-ipafont
RUN apt-get install -y default-jdk 
RUN apt-get install git

RUN mkdir -p /opt/plantuml/ 
RUN wget https://github.com/plantuml/plantuml/releases/download/v1.2024.4/plantuml-mit-1.2024.4.jar -P /opt/plantuml/ 
RUN mv /opt/plantuml/plantuml-mit-1.2024.4.jar /opt/plantuml/plantuml.jar 
RUN chmod a+x /opt/plantuml/plantuml.jar 

WORKDIR /workdir