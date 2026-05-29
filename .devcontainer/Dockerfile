ARG VERSION="slim-trixie"
FROM python:${VERSION}

ARG DEBIAN_FRONTEND="noninteractive"

ENV PYTHONUNBUFFERED=1 \
  TZ="Asia/Tokyo" \
  LANG="ja_JP.UTF-8"

# Install deb packages.
RUN apt-get update && apt-get install -y --no-install-recommends \
  curl \
  fonts-noto-cjk \
  git \
  iproute2 \
  iputils-ping \
  jq \
  locales \
  openssh-client \
  procps \
  rsync \
  sudo \
  unzip zip \
  wget \
  && apt-get clean \
  && rm -fr /var/lib/apt/lists/* \
  && echo 'ja_JP.UTF-8 UTF-8' >/etc/locale.gen \
  && locale-gen \
  && update-locale LANG=ja_JP.UTF-8 \
  && ln -sf /usr/share/zoneinfo/Asia/Tokyo /etc/localtime \
  && rm -fr /var/log/*

ARG OS_APP_GROUP=vscode
ARG LOCAL_GID
ARG OS_APP_USER=vscode
ARG LOCAL_UID

# Create user.
# Configure sudoers.
RUN groupadd -g ${LOCAL_GID:?} ${OS_APP_GROUP} \
  && useradd -l -m -u ${LOCAL_UID:?} -g ${OS_APP_GROUP} -s /bin/bash ${OS_APP_USER} \
  && echo "${OS_APP_USER} ALL=(ALL) NOPASSWD: ALL" >"/etc/sudoers.d/${OS_APP_USER}"

USER ${OS_APP_USER}

RUN mkdir -m 0700 -pv ~/.ssh

# Install pip packages.
RUN --mount=type=bind,target=src \
  pip install --no-cache-dir --upgrade pip \
  && pip install --no-cache-dir -r src/requirements.txt
