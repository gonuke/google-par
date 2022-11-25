FROM ubuntu:22.04

RUN apt-get -y --force-yes update
RUN apt-get install -y pip vim

RUN pip install --upgrade pip

RUN pip install gspread oauth2client

