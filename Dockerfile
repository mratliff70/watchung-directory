# Taken from https://hub.docker.com/r/garland/aws-cli-docker/~/dockerfile/

# To build for Linux
FROM surnet/alpine-python-wkhtmltopdf:3.11.4-0.12.6-small

# To build image for Mac M1 chip architecture
#FROM surnet/alpine-python-wkhtmltopdf:3.11.4-0.12.6-small@sha256:3162e6df62d102daadffe43847b600065e4059dea4e9dc446bb81cea8e62cec7

RUN apk --no-cache update && \
    # Upgrade pip
    pip install --upgrade pip setuptools && \
    # Install required Python modules
    pip --no-cache-dir install PyPDF2 && \
    pip --no-cache-dir install pdfkit && \
    #pip --no-cache-dir install httplib2  && \
    #pip --no-cache-dir install oauth2client  && \
    pip --no-cache-dir install google-api-python-client && \
    pip --no-cache-dir install google-auth-httplib2 && \
    pip --no-cache-dir install google-auth-oauthlib &&\
    pip --no-cache-dir install boto3
    #pip --no-cache-dir install apiclient && \

# Make the directory where our script will be stored

RUN mkdir /data

WORKDIR /data

ADD makedirectory.py /data/

ENTRYPOINT ["python", "makedirectory.py"]
