# Excel to MySQL comparison

`excel-mysql-compare` compares worksheet rows with a MySQL table using one or
more identity columns. Worksheet headers must match MySQL column names. Every
worksheet column is compared; additional MySQL columns are ignored. Blank
worksheet rows are skipped, and duplicate or empty identities are reported as
errors.

## Install

Use Python 3.9 or newer:

```sh
python -m pip install .
```

## Run

Provide database credentials through environment variables. The password is
read only from `MYSQL_PASSWORD` and is never accepted as a command-line
argument.

```sh
export MYSQL_HOST=127.0.0.1
export MYSQL_PORT=3306
export MYSQL_DATABASE=app_db
export MYSQL_USER=readonly_user
export MYSQL_PASSWORD='...'

excel-mysql-compare \
  --excel ./customers.xlsx \
  --sheet Customers \
  --table customers \
  --id-columns customer_id
```

For compound identities, separate column names with commas:

```sh
excel-mysql-compare --excel ./orders.xlsx --sheet Export \
  --table sales.orders --id-columns tenant_id,order_id --output json
```

`--mysql-host`, `--mysql-port`, `--mysql-database`, and `--mysql-user` can
override their respective environment variables. The tool issues a `SELECT`
only in an explicitly read-only transaction, safely quotes table and column
identifiers, and connects without autocommit. Reports include exact matches,
each differing column and its two values, and full rows that appear on only
one side. JSON output is available with `--output json`.

Run the focused tests with:

```sh
python -m unittest discover -s tests
```
