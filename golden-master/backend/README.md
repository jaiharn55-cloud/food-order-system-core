# Golden Master Backend V1.0

Production source of truth: `jaiharn55-cloud/food-order-liff` at commit `7ceb6c1a83fe7d6db7f27352d864e41d86840a92`.

The complete Store-aware backend file is prepared locally as `CORE_BACKEND_MASTER_FULL_STORE_AWARE.gs`.

It retains the production function set and removes hardcoded production values for:
- LIFF ID
- shop name
- GitHub owner/repository/branch/image directory
- production Admin PIN

Store-specific values are supplied by Apps Script Script Properties / `STORE_CONFIG_JSON`.

Secrets such as `GITHUB_TOKEN`, `LINE_CHANNEL_ACCESS_TOKEN`, and `SHOP_NOTIFY_GROUP_ID` stay in Script Properties and must not be committed here.

Before treating this directory as the Golden Master, upload `CORE_BACKEND_MASTER_FULL_STORE_AWARE.gs` as `CORE_BACKEND_MASTER.gs` in this directory. The file has been syntax-checked as JavaScript source before delivery.
