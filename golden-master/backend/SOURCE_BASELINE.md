# Backend Golden Master Source Baseline

The live production backend is the current source of truth for backend functionality.

- Production repo: `jaiharn55-cloud/food-order-liff`
- Production baseline commit: `7ceb6c1a83fe7d6db7f27352d864e41d86840a92`
- Production backend source is maintained in the Google Apps Script project used by the live store.
- The current baseline contains 59 named functions.

## Rule

Do not replace the backend with a reduced rewrite. New Store Instances must retain every current production capability, then receive configuration for the target store.

## Store-specific runtime settings

These belong in Script Properties / secure configuration rather than source code:

- `SHOP_NAME`
- `STORE_ID`
- `LIFF_ID`
- `LIFF_URL`
- `API_URL`
- `GITHUB_OWNER`
- `GITHUB_REPO`
- `GITHUB_BRANCH`
- `GITHUB_IMAGE_DIR`
- `INITIAL_ADMIN_PIN` / generated admin PIN
- `ADMIN_AUTH_PIN`
- `LINE_CHANNEL_ACCESS_TOKEN`

The live code already reads the LINE channel access token and GitHub token from Script Properties; secrets must remain out of Git and onboarding Forms.
