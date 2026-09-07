# FOOD ORDER SYSTEM CORE

Core / Golden Master สำหรับระบบสั่งอาหารผ่าน LINE ที่จะนำไปสร้าง Store Instance ให้ร้านใหม่หลายร้าน

## Golden Master source
- Production repository: `jaiharn55-cloud/food-order-liff`
- Source commit: `7ceb6c1`
- Production repository ห้ามถูกแก้โดย Store Builder

## Current capability baseline
Customer:
- LINE LIFF customer ordering
- Menu / category / special menu
- Menu options and option pricing
- Daily curry / rice-with-curry selection
- Cart / quantity / remarks
- Pickup slots / capacity / queue waiting count
- Cash / PromptPay / government payment modes
- My Orders / order status

Admin:
- Admin login with PIN / session
- Order management and status updates
- Payment status management
- Menu CRUD / availability
- Daily Dish Library / Daily Dishes
- Special Menus
- Payment settings / QR
- Menu image upload to GitHub
- Order-ready LINE notification

Backend / data:
- Google Apps Script Web App API
- Google Sheets data model
- Orders + Order_Detail + Customers
- Pickup slot management
- Store configuration

## Architecture

`CORE CODE + STORE CONFIG + STORE DATABASE = STORE INSTANCE`

Store Builder must create isolated instances. A change to one Store Instance must not modify another Store Instance or Production.

## Important rule

Credentials and secrets are not stored in this repository or Google Forms. Use Script Properties / GitHub secrets / LINE credentials management as appropriate.
