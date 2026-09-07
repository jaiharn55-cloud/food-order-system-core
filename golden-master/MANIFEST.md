# Golden Master V1.0

## Production baseline
- Repository: `jaiharn55-cloud/food-order-liff`
- Baseline commit: `7ceb6c1a83fe7d6db7f27352d864e41d86840a92`
- Customer: `index.html`
- Admin: `admin.html`

The production repository is the live reference and must not be modified by Store Builder.

## Frontend source blobs
- Customer `index.html` blob: `04359ee33164d8cbb967ad403e7236d0eeed9e93`
- Admin `admin.html` blob: `10c127b8c8f955b38af74ed6e926865360c951b1`

## Backend baseline
The production backend baseline contains 59 named functions. The Store Core will retain all production capabilities; store-specific values are moved to per-instance configuration / Script Properties rather than being deleted.

## Store-specific separation
- SHOP_NAME
- STORE_ID
- LIFF_ID / LIFF_URL
- API URL
- LINE channel identifiers
- GitHub repository / branch / image directory
- Payment settings
- Pickup settings
- Admin PIN and sessions

Secrets such as LINE channel access token and GitHub token must never be stored in this repository or in the customer onboarding form.
