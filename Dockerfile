# Taken from https://hub.docker.com/r/garland/aws-cli-docker/~/dockerfile/

FROM surnet/alpine-python-wkhtmltopdf:3.6.5-0.12.5-small

RUN apk --no-cache update && \
    # Upgrade pip
    pip install --upgrade pip setuptools && \
    # Install required Python modules
    pip --no-cache-dir install PyPDF2 && \
    pip --no-cache-dir install pdfkit && \
    pip --no-cache-dir install httplib2  && \
    pip --no-cache-dir install oauth2client  && \
    pip --no-cache-dir install google-api-python-client
    pip --no-cache-dir install boto3 && \
    #pip --no-cache-dir install apiclient && \

WORKDIR /data

ADD makedirectory.py /data/

ADD token.json /data/

ADD credentials.json /data/
