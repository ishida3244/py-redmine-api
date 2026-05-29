# redmine_api

Python 開発環境のコンテナ

<!-- TOC -->

- [1. Usage](#1-usage)
  - [1.1. Redmine REST API の使用時](#11-redmine-rest-api-の使用時)
- [2. Author](#2-author)

<!-- /TOC -->


## 1. Usage

### 1.1. Redmine REST API の使用時

- `/workspace/.env.local` を作成して以下の例のように環境変数を定義する
- `API_KEY` の値は Redmine の個人設定 ＞ APIアクセスキー で作成する
- `API_URL` の値は `プロトコル://ホスト[:ポート]`

```shell
API_KEY=0123456789abcdef0123456789abcdef01234567
API_URL=http://redmine.wk.aruze.co.jp:3000
```

```shell
API_KEY=0123456789abcdef0123456789abcdef01234567
API_URL=http://localhost
```

## 2. Author

ishida
