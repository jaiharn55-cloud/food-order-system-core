/** FOOD ORDER SYSTEM CORE BACKEND V1.0 - GOLDEN MASTER
 * Derived from production baseline commit 7ceb6c1.
 * 63 functions retained; store-specific values are configurable.
 */

/**
 * ============================================================
 * STORE CONFIGURATION ADAPTER
 * ============================================================
 * Core code is store-neutral. Each Store Instance supplies
 * store-specific values through Apps Script Script Properties.
 *
 * Required / optional keys:
 *   STORE_ID
 *   SHOP_NAME
 *   LIFF_ID
 *   GITHUB_OWNER
 *   GITHUB_REPO
 *   GITHUB_BRANCH
 *   GITHUB_IMAGE_DIR
 *
 * Secrets remain in Script Properties and are never placed in
 * this source repository or in the customer onboarding form.
 * ============================================================
 */
function getStoreProperty_(key, fallback) {
  const value = String(
    PropertiesService
      .getScriptProperties()
      .getProperty(key) || ''
  ).trim();

  return value || (fallback === undefined ? '' : fallback);
}

const CONFIG = {
  LIFF_ID: getStoreProperty_('LIFF_ID', ''),
  STORE_ID: getStoreProperty_('STORE_ID', ''),
  SHOP_NAME: getStoreProperty_('SHOP_NAME', 'FOOD ORDER SYSTEM'),

  SHEETS: {
    MENU: 'Menu',
    OPTIONS: 'Menu_Options',
    DAILY_DISH_LIBRARY: 'Daily_Dish_Library',
    ORDERS: 'Orders',
    DETAILS: 'Order_Detail',
    CUSTOMERS: 'Customers',
    SLOTS: 'Pickup_Slots'
  }
};


/**
 * เปิดหน้าเว็บ
 */
function doGet(e) {

  // API สำหรับโหลดข้อมูล
  if (
    e &&
    e.parameter &&
    e.parameter.action === 'getInitialData'
  ) {

    return jsonResponse_(
      getInitialData()
    );

  }

  if (
  e &&
  e.parameter &&
  e.parameter.action ===
  'getOrders'
) {

  
    requireAdminAuth_(
      e.parameter.adminToken
    );

return jsonResponse_(
    getOrders()
  );

}


  // ================================
  // CUSTOMER - MY ORDERS
  // ================================
  if (
    e &&
    e.parameter &&
    e.parameter.action ===
    'getMyOrders'
  ) {

    return jsonResponse_(
      getMyOrders(
        e.parameter.userId
      )
    );

  }

  if (
    e.parameter.action ===
    'getDailyDishLibrary'
  ) {

    requireAdminAuth_(
      e.parameter.adminToken
    );

    return jsonResponse_(
      getDailyDishLibrary()
    );

  }

  // ================================
  // MENU SYSTEM
  // ================================
  if (
    e &&
    e.parameter &&
    e.parameter.action ===
    'getMenuSystemData'
  ) {

    
    requireAdminAuth_(
      e.parameter.adminToken
    );

return jsonResponse_(
      getMenuSystemData()
    );

  }

  // ================================
  // ADMIN SETTINGS
  // ================================
  if (
    e &&
    e.parameter &&
    e.parameter.action ===
    'getAdminSettings'
  ) {

    requireAdminAuth_(
      e.parameter.adminToken
    );

    return jsonResponse_(
      getAdminSettings()
    );

  }

  // หน้าเว็บเดิมของ Apps Script
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.SHOP_NAME || 'FOOD ORDER SYSTEM')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/**
 * โหลดข้อมูลเริ่มต้น
 */
function getInitialData() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const menus = getSheetObjects_(ss.getSheetByName(CONFIG.SHEETS.MENU));
  const options = getSheetObjects_(ss.getSheetByName(CONFIG.SHEETS.OPTIONS));
  const slots = getSheetObjects_(ss.getSheetByName(CONFIG.SHEETS.SLOTS));
  const orders = getSheetObjects_(ss.getSheetByName(CONFIG.SHEETS.ORDERS));

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );

  const slotCounts = {};
  const waitingCounts = {};

  // คิวที่รอ = เฉพาะออเดอร์สถานะ NEW + PREPARING
  orders.forEach(function(order) {
    let orderDate = '';
    const rawDate = order.Date;
    if (
      Object.prototype.toString.call(rawDate) === '[object Date]' &&
      !isNaN(rawDate.getTime())
    ) {
      orderDate = Utilities.formatDate(
        rawDate,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      );
    } else {
      orderDate = String(rawDate || '').trim();
    }

    if (orderDate === today) {
      const slot = String(order.Pickup_Slot || '').trim();
      const status = String(order.Status || '').trim().toUpperCase();
      if (slot) {
        if (status !== 'COMPLETED' && status !== 'CANCELLED') {
          slotCounts[slot] = (slotCounts[slot] || 0) + 1;
        }
        if (status === 'NEW' || status === 'PREPARING') {
          waitingCounts[slot] = (waitingCounts[slot] || 0) + 1;
        }
      }
    }
  });

  const dailySheet = ss.getSheetByName('Daily_Dishes');
  const dailyRows = dailySheet ? getSheetObjects_(dailySheet) : [];
  const dailyDishes = dailyRows
  .filter(function(dish) {
    return normalizeDateValue_(dish.Dish_Date) === today &&
      String(dish.Available || '').toUpperCase() === 'TRUE';
  })
    .map(function(dish) {
      return {
        id: String(dish.Dish_ID || ''),
        name: String(dish.Dish_Name || ''),
        description: String(dish.Description || ''),
        image: String(dish.Image_URL || '')
      };
    });

  const specialSheet = ss.getSheetByName('Special_Menus');
  const specialRows = specialSheet ? getSheetObjects_(specialSheet) : [];
  const specialMenus = specialRows
    .filter(function(menu) {
      const start = normalizeDateValue_(
      menu.Start_Date || today
    );

    const end = normalizeDateValue_(
      menu.End_Date || today
    );

    return String(menu.Available || '').toUpperCase() === 'TRUE' &&
      today >= start &&
      today <= end;
    })
    .map(function(menu) {
      return {
        id: String(menu.Special_ID || ''),
        name: String(menu.Menu_Name || ''),
        description: String(menu.Description || ''),
        price: Number(menu.Price || 0),
        category: '⭐ เมนูพิเศษ',
        image: String(menu.Image_URL || ''),
        menuType: 'SPECIAL'
      };
    });

  const mainMenus = menus
    .filter(function(menu) {
      return String(menu.Available || '').toUpperCase() === 'TRUE';
    })
    .map(function(menu) {
      return {
        id: String(menu.Menu_ID || ''),
        name: String(menu.Menu_Name || ''),
        description: String(menu.Description || ''),
        price: Number(menu.Price || 0),
        category: String(menu.Category || ''),
        image: String(menu.Image_URL || ''),
        menuType: 'MAIN'
      };
    });

  const visibleOptions = options
    .filter(function(option) {
      return String(option.Available || '').toUpperCase() === 'TRUE';
    })
    .map(function(option) {
      return {
        menuId: String(option.Menu_ID || ''),
        name: String(option.Option_Name || ''),
        price: Number(option.Option_Price || 0),
        group: String(option.Option_Group || ''),
        minSelect: Number(option.Min_Select || 0),
        maxSelect: Number(option.Max_Select || 0)
      };
    });

  const curryConfig = {
    M022: 1,
    M023: 2,
    M024: 3
  };

  Object.keys(curryConfig).forEach(function(menuId) {
    const required = curryConfig[menuId];
    dailyDishes.forEach(function(dish) {
      visibleOptions.push({
        menuId: menuId,
        name: dish.name,
        price: 0,
        group: 'เลือกกับข้าว',
        minSelect: required,
        maxSelect: required,
        dynamicDailyDish: true
      });
    });
  });

  return {
    appVersion: 'BACKEND-V4-MENU-SYSTEM',
    liffId: CONFIG.LIFF_ID,
    payment: getPaymentSettings_(),
    menus: mainMenus.concat(specialMenus),
    options: visibleOptions,
    dailyDishes: dailyDishes,
    specialMenus: specialMenus,
    slots: slots
      .filter(function(slot) {
        return String(slot.Active || '').toUpperCase() === 'TRUE';
      })
      .map(function(slot) {
        const id = String(slot.Slot_ID || '');
        const capacity = Number(slot.Capacity || 0);
        const used = slotCounts[id] || 0;
        return {
          id: id,
          start: formatTime_(slot.Start_Time),
          end: formatTime_(slot.End_Time),
          capacity: capacity,
          used: used,
          waiting: waitingCounts[id] || 0,
          waitingCount: waitingCounts[id] || 0,
          remaining: Math.max(capacity - used, 0)
        };
      })
  };
}


/**
 * สร้าง Order
 */
function createOrder(orderData) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(15000);


  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const orderSheet =
      ss.getSheetByName(
        CONFIG.SHEETS.ORDERS
      );

    const detailSheet =
      ss.getSheetByName(
        CONFIG.SHEETS.DETAILS
      );

    const customerSheet =
      ss.getSheetByName(
        CONFIG.SHEETS.CUSTOMERS
      );

    const menuSheet =
      ss.getSheetByName(
        CONFIG.SHEETS.MENU
      );

    const optionSheet =
      ss.getSheetByName(
        CONFIG.SHEETS.OPTIONS
      );

    const slotSheet =
      ss.getSheetByName(
        CONFIG.SHEETS.SLOTS
      );


    if (!orderData) {
      throw new Error(
        'ไม่พบข้อมูลออเดอร์'
      );
    }


    if (
      !orderData.items ||
      orderData.items.length === 0
    ) {
      throw new Error(
        'ยังไม่มีอาหารในตะกร้า'
      );
    }


    if (!orderData.userId) {
      throw new Error(
        'ไม่พบ LINE User ID'
      );
    }


    const menus = getSheetObjects_(menuSheet);
    const options = getSheetObjects_(optionSheet);

    const dailySheet = ss.getSheetByName('Daily_Dishes');

    const menuSystemToday = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

    const dailyDishes = dailySheet
      ? getSheetObjects_(dailySheet).filter(function(dish) {
          return normalizeDateValue_(dish.Dish_Date) === menuSystemToday &&
            String(dish.Available || '').toUpperCase() === 'TRUE';
        })
      : [];

    const specialSheet = ss.getSheetByName('Special_Menus');
    const specialMenus = specialSheet
      ? getSheetObjects_(specialSheet).filter(function(menu) {
          const startDate = String(menu.Start_Date || menuSystemToday).trim();
          const endDate = String(menu.End_Date || menuSystemToday).trim();
          return String(menu.Available || '').toUpperCase() === 'TRUE' &&
            menuSystemToday >= startDate && menuSystemToday <= endDate;
        })
      : [];

    const menuMap = {};

    menus.forEach(function(menu) {
      if (String(menu.Available || '').toUpperCase() === 'TRUE') {
        menuMap[String(menu.Menu_ID)] = {
          id: String(menu.Menu_ID),
          name: String(menu.Menu_Name),
          price: Number(menu.Price || 0)
        };
      }
    });

    specialMenus.forEach(function(menu) {
      menuMap[String(menu.Special_ID)] = {
        id: String(menu.Special_ID),
        name: String(menu.Menu_Name),
        price: Number(menu.Price || 0)
      };
    });

    let total = 0;

    const detailRows = [];


    orderData.items.forEach(
      function(item) {

        const menu =
          menuMap[
            String(item.menuId)
          ];


        if (!menu) {

          throw new Error(
            'ไม่พบเมนู: ' +
            item.menuId
          );

        }


        const qty =
          Math.max(
            1,
            Number(item.qty || 1)
          );


        let optionTotal = 0;

        const optionNames = [];


                 const curryRequiredMap = {
           M022: 1,
           M023: 2,
           M024: 3
         };

         const curryRequired =
           curryRequiredMap[String(item.menuId)] || 0;

         if (curryRequired > 0) {
           const selectedNames = (item.options || [])
             .map(function(selected) {
               return String(selected.name || '').trim();
             })
             .filter(Boolean);

           const uniqueNames = selectedNames.filter(function(name, index, arr) {
             return arr.indexOf(name) === index;
           });

           if (uniqueNames.length !== curryRequired) {
             throw new Error(
               'กรุณาเลือกกับข้าว ' + curryRequired + ' อย่าง'
             );
           }
         }

(item.options || [])
          .forEach(function(selected) {

            const found =
              options.find(function(option) {
                return String(option.Menu_ID) === String(item.menuId) &&
                  String(option.Option_Name) === String(selected.name) &&
                  String(option.Available).toUpperCase() === 'TRUE';
              });

            const isDailyDishOption =
              ['M022', 'M023', 'M024'].indexOf(String(item.menuId)) >= 0 &&
              dailyDishes.some(function(dish) {
                return String(dish.Dish_Name) === String(selected.name);
              });

            if (found || isDailyDishOption) {
              optionTotal += found ? Number(found.Option_Price || 0) : 0;
              optionNames.push(String(selected.name));
            }

          });


        const itemTotal =
          (
            menu.price +
            optionTotal
          ) * qty;


        total += itemTotal;


        detailRows.push([

          Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            'yyyy-MM-dd'
          ),

          '',
          menu.id,
          menu.name,
          qty,
          menu.price,
          optionNames.join(', '),
          optionTotal,
          itemTotal,
          String(item.remark || '')

        ]);

      }
    );


    const pickupType =
      String(
        orderData.pickupType || 'NOW'
      );


    let pickupSlot = '';


    if (
      pickupType === 'SCHEDULED'
    ) {

      pickupSlot =
        String(
          orderData.pickupSlot || ''
        );


      if (!pickupSlot) {

        throw new Error(
          'กรุณาเลือกช่วงเวลารับ'
        );

      }


      const slots =
        getSheetObjects_(slotSheet);


      const selectedSlot =
        slots.find(
          function(slot) {

            return (
              String(slot.Slot_ID) ===
                pickupSlot &&

              String(slot.Active)
                .toUpperCase() ===
                'TRUE'
            );

          }
        );


      if (!selectedSlot) {

        throw new Error(
          'ช่วงเวลานี้ไม่สามารถเลือกได้'
        );

      }


      const today =
        Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        );


      const orders =
  getSheetObjects_(orderSheet);

let used = 0;


orders.forEach(
  function(order) {

    let orderDate = '';

    const rawDate =
      order.Date;


    if (
      Object.prototype.toString
        .call(rawDate) ===
      '[object Date]' &&
      !isNaN(rawDate.getTime())
    ) {

      orderDate =
        Utilities.formatDate(
          rawDate,
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        );

    } else {

      orderDate =
        String(
          rawDate || ''
        ).trim();

    }


    const status =
      String(
        order.Status || ''
      )
      .trim()
      .toUpperCase();


    const orderSlot =
      String(
        order.Pickup_Slot || ''
      ).trim();


    if (
      orderDate === today &&
      orderSlot === pickupSlot &&
      status !== 'COMPLETED' &&
      status !== 'CANCELLED'
    ) {

      used++;

    }

  }
);


      const capacity =
        Number(
          selectedSlot.Capacity || 0
        );


      if (used >= capacity) {

        throw new Error(
          'ช่วงเวลานี้เต็มแล้ว กรุณาเลือกเวลาอื่น'
        );

      }

    }


    // ==========================================
// สร้างเลข Order ID ของวันนี้
// ==========================================

const today =
  Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );


const lastRow =
  orderSheet.getLastRow();


let maxNumber = 0;


if (lastRow >= 2) {

  // อ่านคอลัมน์ A = Order_ID
  // และ B = Date โดยตรงจาก Sheet
  const values =
    orderSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getValues();


  values.forEach(
    function(row) {

      const rawOrderId =
        row[0];


      const rawDate =
        row[1];


      // แปลงวันที่จาก Google Sheets
      let rowDate = '';


      if (
        Object.prototype.toString
          .call(rawDate) ===
        '[object Date]' &&
        !isNaN(rawDate.getTime())
      ) {

        rowDate =
          Utilities.formatDate(
            rawDate,
            Session.getScriptTimeZone(),
            'yyyy-MM-dd'
          );

      } else {

        rowDate =
          String(
            rawDate || ''
          ).trim();

      }


      // ต้องเป็นออเดอร์ของวันนี้เท่านั้น
      if (
        rowDate !== today
      ) {

        return;

      }


      const id =
        String(
          rawOrderId || ''
        )
        .trim()
        .toUpperCase();


      const match =
        id.match(/^A(\d+)$/);


      if (match) {

        const number =
          Number(match[1]);


        if (
          !isNaN(number) &&
          number > maxNumber
        ) {

          maxNumber =
            number;

        }

      }

    }
  );

}


const orderId =
  'A' +
  String(
    maxNumber + 1
  ).padStart(3, '0');


// Debug สำหรับตรวจสอบ
console.log(
  'Today:',
  today
);

console.log(
  'Max Order Number:',
  maxNumber
);

console.log(
  'New Order ID:',
  orderId
);


    const now = new Date();


    const date =
      Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      );


    const time =
      Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        'HH:mm:ss'
      );


    const createdAt =
      Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd HH:mm:ss'
      );


    const payment =
      String(
        orderData.payment || 'CASH'
      );


    const customerName =
      String(
        orderData.customerName || ''
      );


    const customerRemark =
      String(
        orderData.customerRemark || ''
      );


    orderSheet.appendRow([

      orderId,
      date,
      time,
      String(orderData.userId),
      customerName,
      pickupType,
      pickupSlot,
      total,
      payment,
      'UNPAID',
      'NEW',
      customerRemark,
      createdAt

    ]);


    detailRows.forEach(
      function(row) {

        row[1] = orderId;

        detailSheet.appendRow(row);

      }
    );


        saveCustomer_(
          customerSheet,
          String(orderData.userId),
          customerName
        );

        /*
        * ==========================================
        * PROD - แจ้งร้านเมื่อมีออเดอร์ใหม่
        *
        * Fail-safe:
        * LINE มีปัญหา
        * แต่ Order ยังต้องสำเร็จ
        * ==========================================
        */
        let shopNotification = {
          success: false,
          sent: false,
          reason:
            'NOT_ATTEMPTED'
        };

        try {

          shopNotification =
            sendNewOrderGroupNotificationProd_(
              orderId,
              orderData,
              total,
              detailRows,
              date,
              time
            );

        } catch (error) {

          console.error(
            'PROD SHOP NEW ORDER NOTIFICATION ERROR:',
            error
          );

          shopNotification = {
            success: false,
            sent: false,
            reason:
              error.message ||
              String(error)
          };

        }

        return {
          success: true,

          orderId:
            orderId,

          total:
            total,

          pickupType:
            pickupType,

          pickupSlot:
            pickupSlot,

          shopNotification:
            shopNotification
        };


  } finally {

    lock.releaseLock();

  }

}


/**
 * บันทึกลูกค้า
 */
function saveCustomer_(
  sheet,
  userId,
  name
) {

  if (!sheet) return;


  const data =
    getSheetObjects_(sheet);


  const index =
    data.findIndex(
      function(row) {

        return (
          String(
            row.LINE_User_ID
          ) === userId
        );

      }
    );


  const now = new Date();


  const lastOrder =
    Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );


  if (index >= 0) {

    const rowNumber =
      index + 2;


    const headers =
      getHeaders_(sheet);


    const nameCol =
      headers.indexOf(
        'Name'
      ) + 1;


    const displayCol =
      headers.indexOf(
        'Display_Name'
      ) + 1;


    const lastOrderCol =
      headers.indexOf(
        'Last_Order'
      ) + 1;


    if (nameCol > 0) {

      sheet
        .getRange(
          rowNumber,
          nameCol
        )
        .setValue(name);

    }


    if (displayCol > 0) {

      sheet
        .getRange(
          rowNumber,
          displayCol
        )
        .setValue(name);

    }


    if (lastOrderCol > 0) {

      sheet
        .getRange(
          rowNumber,
          lastOrderCol
        )
        .setValue(lastOrder);

    }


  } else {

    sheet.appendRow([

      userId,
      name,
      name,
      lastOrder

    ]);

  }

}


/**
 * อ่าน Sheet
 */
function getSheetObjects_(sheet) {

  if (!sheet) {

    throw new Error(
      'ไม่พบ Sheet ที่ต้องการ'
    );

  }


  const values =
    sheet
      .getDataRange()
      .getValues();


  if (values.length < 2) {
    return [];
  }


  const headers =
    values[0].map(
      function(header) {

        return String(header)
          .trim();

      }
    );


  return values
    .slice(1)
    .filter(
      function(row) {

        return row.some(
          function(cell) {

            return cell !== '';

          }
        );

      }
    )
    .map(
      function(row) {

        const obj = {};


        headers.forEach(
          function(header, index) {

            obj[header] =
              row[index];

          }
        );


        return obj;

      }
    );

}

/**
 * แปลงวันที่จาก Google Sheets / Text
 * ให้เป็น yyyy-MM-dd เหมือนกันทั้งระบบ
 */
function normalizeDateValue_(value) {

  if (
    Object.prototype.toString.call(value) === '[object Date]' &&
    !isNaN(value.getTime())
  ) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  const text = String(value || '').trim();

  // รองรับ ISO datetime
  // เช่น 2026-08-24T00:00:00.000Z
  if (/^\d{4}-\d{2}-\d{2}T/.test(text)) {
    return text.substring(0, 10);
  }

  return text;
}

/**
 * อ่านหัวตาราง
 */
function getHeaders_(sheet) {

  return sheet
    .getRange(
      1,
      1,
      1,
      sheet.getLastColumn()
    )
    .getValues()[0]
    .map(
      function(header) {

        return String(header)
          .trim();

      }
    );

}

/**
 * แปลงเวลา Google Sheets ให้เป็น HH:mm
 */
function formatTime_(value) {

  // Google Sheets ส่งเวลาเป็น Date object
  if (Object.prototype.toString.call(value) === '[object Date]') {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'HH:mm'
    );

  }

  // กรณี Google Sheets ส่งเป็นตัวเลข
  if (
    typeof value === 'number' &&
    !isNaN(value)
  ) {

    const totalMinutes =
      Math.round(value * 24 * 60);

    const hours =
      Math.floor(totalMinutes / 60) % 24;

    const minutes =
      totalMinutes % 60;

    return (
      String(hours).padStart(2, '0') +
      ':' +
      String(minutes).padStart(2, '0')
    );

  }

  // กรณีเป็นข้อความ
  if (
    value === null ||
    value === undefined
  ) {

    return '';

  }

  const text =
    String(value).trim();

  const match =
    text.match(/(\d{1,2}):(\d{2})/);

  if (match) {

    return (
      String(match[1]).padStart(2, '0') +
      ':' +
      match[2]
    );

  }

  return text;

}

function testPickupSlots() {

  const data = getInitialData();

  Logger.log(
    JSON.stringify(data.slots)
  );

}

function doGetDailyDishLibrary_(e) {

  requireAdminAuth_(
    e.parameter.adminToken
  );

  return jsonResponse_(
    getDailyDishLibrary()
  );

}

function doPost(e) {

  try {

    if (!e || !e.postData) {

      return jsonResponse_({

        success: false,

        error:
          'ไม่พบข้อมูลที่ส่งมา'

      });

    }


    const data =
      JSON.parse(
        e.postData.contents
      );


    const action =
      String(
        data.action || ''
      );


    // ================================
    // สร้าง Order
    // ================================

    if (
      action === 'createOrder'
    ) {

      const result =
        createOrder(
          data.orderData
        );


      return jsonResponse_(
        result
      );

    }

    if (
  action ===
  'updateOrderStatus'
) {

  
      requireAdminAuth_(
        data.adminToken
      );

const result =
    updateOrderStatus(
      data.orderId,
      data.orderDate,
      data.status
);


  return jsonResponse_(
    result
  );

}


/* PAYMENT STATUS: added without changing CONFIG/QR/order logic */
if (
  action ===
  'updatePaymentStatus'
) {

  
      requireAdminAuth_(
        data.adminToken
      );

const result =
    updatePaymentStatus(
      data.orderId,
      data.orderDate,
      data.paymentStatus
    );

  return jsonResponse_(
    result
  );

}


    
    // ================================
    // ADMIN AUTH
    // ================================

    if (
      action === 'loginAdmin'
    ) {

      return jsonResponse_(
        loginAdmin(
          data.pin
        )
      );

    }


    if (
      action === 'logoutAdmin'
    ) {

      return jsonResponse_(
        logoutAdmin(
          data.token
        )
      );

    }


    if (
      action === 'changeAdminPin'
    ) {

      return jsonResponse_(
        changeAdminPin(
          data.currentPin,
          data.newPin,
          data.token
        )
      );

    }


    // ================================
    // MENU SYSTEM V2
    // ================================

    if (action === 'installRecommendedMenuSystem') {
      requireAdminAuth_(data.adminToken);

      return jsonResponse_(installRecommendedMenuSystem(data.confirm));
    }

    if (
      action ===
      'addDailyDishesFromLibrary'
    ) {

      requireAdminAuth_(
        data.adminToken
      );

      return jsonResponse_(
        addDailyDishesFromLibrary(
          data.date,
          data.libraryIds
        )
      );

    }

    if (action === 'createDailyDish') {
      requireAdminAuth_(data.adminToken);

      return jsonResponse_(createDailyDish(data.dish));
    }

    if (action === 'saveDailyDish') {
      requireAdminAuth_(data.adminToken);

      return jsonResponse_(saveDailyDish(data.dish));
    }

    if (action === 'toggleDailyDish') {
      requireAdminAuth_(data.adminToken);

      return jsonResponse_(toggleDailyDish(data.dishId, data.available));
    }

    if (action === 'deleteDailyDish') {
      requireAdminAuth_(data.adminToken);

      return jsonResponse_(deleteDailyDish(data.dishId));
    }

    if (action === 'createSpecialMenu') {
      requireAdminAuth_(data.adminToken);

      return jsonResponse_(createSpecialMenu(data.menu));
    }

    if (action === 'saveSpecialMenu') {
      requireAdminAuth_(data.adminToken);

      return jsonResponse_(saveSpecialMenu(data.menu));
    }

    if (action === 'toggleSpecialMenu') {
      requireAdminAuth_(data.adminToken);

      return jsonResponse_(toggleSpecialMenu(data.menuId, data.available));
    }

    // ================================
    // ADMIN SETTINGS
    // ================================

    if (
      action === 'saveMenu'
    ) {

      requireAdminAuth_(data.adminToken);

      return jsonResponse_(
        saveMenu(
          data.menu
        )
      );

    }


    if (
      action === 'toggleMenuAvailable'
    ) {

      requireAdminAuth_(data.adminToken);

      return jsonResponse_(
        toggleMenuAvailable(
          data.menuId,
          data.available
        )
      );

    }


    if (
      action === 'createMenu'
    ) {

      requireAdminAuth_(data.adminToken);

      return jsonResponse_(
        createMenu(
          data.menu
        )
      );

    }


    if (
      action === 'savePaymentSettings'
    ) {

      requireAdminAuth_(data.adminToken);

      return jsonResponse_(
        savePaymentSettings(
          data.settings
        )
      );

    }


    if (
      action === 'uploadAdminImage'
    ) {

      requireAdminAuth_(data.adminToken);

      return jsonResponse_(
        uploadAdminImage(
          data.image
        )
      );

    }


return jsonResponse_({

      success: false,

      error:
        'ไม่รู้จัก action: ' +
        action

    });


  } catch (error) {

    return jsonResponse_({

      success: false,

      error:
        error.message ||
        String(error)

    });

  }

}

function jsonResponse_(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}


function getMyOrders(
  userId
) {

  const id =
    String(
      userId || ''
    ).trim();

  if (!id) {
    throw new Error(
      'ไม่พบ LINE User ID'
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const orderSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.ORDERS
    );

  const detailSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.DETAILS
    );

  const slotSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.SLOTS
    );

  if (!orderSheet) {
    throw new Error(
      'ไม่พบ Sheet Orders'
    );
  }

  if (!detailSheet) {
    throw new Error(
      'ไม่พบ Sheet Order_Detail'
    );
  }

  const lastRow =
    orderSheet.getLastRow();

  const lastCol =
    orderSheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastCol < 1
  ) {

    return {
      success: true,
      orders: []
    };

  }

  const headers =
    orderSheet
      .getRange(
        1,
        1,
        1,
        lastCol
      )
      .getValues()[0]
      .map(
        function(header) {
          return String(
            header || ''
          ).trim();
        }
      );

  // Current Orders layout stores LINE User ID in column D.
  const userIdHeaderNames = [
    'LINE_User_ID',
    'User_ID',
    'Customer_User_ID',
    'Customer_ID'
  ];

  let userCol = -1;

  userIdHeaderNames.forEach(
    function(name) {

      if (userCol > 0) return;

      const index =
        headers.indexOf(
          name
        );

      if (index >= 0) {
        userCol = index + 1;
      }

    }
  );

  if (
    userCol < 0 &&
    lastCol >= 4
  ) {
    userCol = 4;
  }

  if (userCol < 1) {
    throw new Error(
      'ไม่พบคอลัมน์ LINE User ID ใน Orders'
    );
  }

  const rows =
    orderSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastCol
      )
      .getValues();

  const details =
    getSheetObjects_(
      detailSheet
    );

  const slots =
    slotSheet
      ? getSheetObjects_(
          slotSheet
        )
      : [];

  const slotMap = {};

  slots.forEach(
    function(slot) {

      const slotId =
        String(
          slot.Slot_ID || ''
        ).trim();

      if (!slotId) return;

      slotMap[slotId] = {
        start:
          formatTime_(
            slot.Start_Time
          ),
        end:
          formatTime_(
            slot.End_Time
          )
      };

    }
  );

  const result = [];

  rows.forEach(
    function(row) {

      const rowUserId =
        String(
          row[userCol - 1] || ''
        ).trim();

      if (
        rowUserId !== id
      ) {
        return;
      }

      const obj = {};

      headers.forEach(
        function(header,index) {
          if (header) {
            obj[header] =
              row[index];
          }
        }
      );

      const orderId =
        String(
          obj.Order_ID || ''
        ).trim();

      if (!orderId) return;

      const rawDate =
        obj.Date;

      let orderDate = '';

      if (
        Object.prototype
          .toString
          .call(rawDate) ===
        '[object Date]' &&
        !isNaN(
          rawDate.getTime()
        )
      ) {

        orderDate =
          Utilities.formatDate(
            rawDate,
            Session.getScriptTimeZone(),
            'yyyy-MM-dd'
          );

      } else {

        orderDate =
          String(
            rawDate || ''
          ).trim();

      }

      const orderDetails =
        details
          .filter(
            function(detail) {

              let detailDate = '';

              const detailRawDate =
                detail.Date;

              if (
                Object.prototype
                  .toString
                  .call(
                    detailRawDate
                  ) ===
                  '[object Date]' &&
                !isNaN(
                  detailRawDate.getTime()
                )
              ) {

                detailDate =
                  Utilities.formatDate(
                    detailRawDate,
                    Session.getScriptTimeZone(),
                    'yyyy-MM-dd'
                  );

              } else {

                detailDate =
                  String(
                    detailRawDate || ''
                  ).trim();

              }

              return (
                String(
                  detail.Order_ID || ''
                ).trim() ===
                orderId &&
                detailDate ===
                orderDate
              );

            }
          )
          .map(
            function(detail) {

              return {
                menuName:
                  String(
                    detail.Menu_Name || ''
                  ),
                qty:
                  Number(
                    detail.Qty || 0
                  ),
                options:
                  String(
                    detail.Options || ''
                  ),
                itemTotal:
                  Number(
                    detail.Item_Total || 0
                  )
              };

            }
          );

      const pickupSlot =
        String(
          obj.Pickup_Slot || ''
        ).trim();

      const pickupTime =
        slotMap[pickupSlot]
          ? (
              slotMap[pickupSlot].start +
              ' - ' +
              slotMap[pickupSlot].end
            )
          : 'รับเร็วที่สุด';

      result.push({

        orderId:
          orderId,

        date:
          orderDate,

        time:
          formatTimeValue_(
            obj.Time
          ),

        customerName:
          String(
            obj.Customer_Name || ''
          ),

        pickupType:
          String(
            obj.Pickup_Type || 'NOW'
          ),

        pickupSlot:
          pickupSlot,

        pickupTime:
          pickupTime,

        total:
          Number(
            obj.Total || 0
          ),

        payment:
          String(
            obj.Payment || ''
          ),

        paymentStatus:
          String(
            obj.Payment_Status || ''
          ),

        status:
          String(
            obj.Status || 'NEW'
          )
          .toUpperCase(),

        remark:
          String(
            obj.Customer_Remark || ''
          ),

        items:
          orderDetails

      });

    }
  );

  result.sort(
    function(a,b) {

      const da =
        (a.date || '') +
        ' ' +
        (a.time || '');

      const db =
        (b.date || '') +
        ' ' +
        (b.time || '');

      return db.localeCompare(
        da
      );

    }
  );

  return {
    success: true,
    orders:
      result.slice(
        0,
        10
      )
  };

}


function getOrders() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const orderSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.ORDERS
    );

  const detailSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.DETAILS
    );

    const slotSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.SLOTS
    );

  if (!orderSheet) {

    throw new Error(
      'ไม่พบ Sheet Orders'
    );

  }


  if (!detailSheet) {

    throw new Error(
      'ไม่พบ Sheet Order_Detail'
    );

  }


  const orders =
    getSheetObjects_(
      orderSheet
    );


  const details =
    getSheetObjects_(
      detailSheet
    );

  const slots =
  getSheetObjects_(
    slotSheet
  );


  const slotMap = {};


  slots.forEach(
    function(slot) {

      const slotId =
        String(
          slot.Slot_ID || ''
        ).trim();


      if (slotId) {

        slotMap[slotId] = {

          start:
            formatTime_(
              slot.Start_Time
            ),

          end:
            formatTime_(
              slot.End_Time
            )

        };

      }

    }
  );


  const today =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );


  // ==============================
  // เอาเฉพาะออเดอร์วันนี้
  // ==============================

  const todayOrders =
    orders.filter(
      function(order) {

        let orderDate = '';

        const rawDate =
          order.Date;


        if (
          Object.prototype.toString
            .call(rawDate) ===
            '[object Date]' &&
          !isNaN(rawDate.getTime())
        ) {

          orderDate =
            Utilities.formatDate(
              rawDate,
              Session.getScriptTimeZone(),
              'yyyy-MM-dd'
            );

        } else {

          orderDate =
            String(
              rawDate || ''
            ).trim();

        }


        return (
          orderDate === today
        );

      }
    );


  // ==============================
  // สร้างข้อมูล Order
  // ==============================

  const result =
    todayOrders.map(
      function(order) {

        const orderId =
          String(
            order.Order_ID || ''
          );


        const orderDetails =
          details.filter(
            function(detail) {

              let detailDate = '';

              const rawDate =
                detail.Date;


              if (
                Object.prototype
                  .toString
                  .call(rawDate) ===
                '[object Date]' &&
                !isNaN(
                  rawDate.getTime()
                )
              ) {

                detailDate =
                  Utilities.formatDate(
                    rawDate,
                    Session.getScriptTimeZone(),
                    'yyyy-MM-dd'
                  );

              } else {

                detailDate =
                  String(
                    rawDate || ''
                  ).trim();

              }


              return (
                String(
                  detail.Order_ID || ''
                ).trim() === orderId &&

                detailDate === today
              );

            }
          );


        return {

          orderId:
            orderId,

          date:
            today,

          time:
            formatTimeValue_(
              order.Time
            ),

          customerName:
            String(
              order.Customer_Name || ''
            ),

          pickupType:
            String(
              order.Pickup_Type || 'NOW'
            ),

          pickupSlot:
            String(
              order.Pickup_Slot || ''
            ),

          pickupTime:
            getPickupTimeText_(
              slotMap,
              order.Pickup_Slot
            ),

          total:
            Number(
              order.Total || 0
            ),

          payment:
            String(
              order.Payment || ''
            ),

          paymentStatus:
            String(
              order.Payment_Status || ''
            ),

          status:
            String(
              order.Status || 'NEW'
            )
            .toUpperCase(),

          remark:
            String(
              order.Customer_Remark || ''
            ),

          items:
            orderDetails.map(
              function(detail) {

                return {

                  menuId:
                    String(
                      detail.Menu_ID || ''
                    ),

                  menuName:
                    String(
                      detail.Menu_Name || ''
                    ),

                  qty:
                    Number(
                      detail.Qty || 0
                    ),

                  basePrice:
                    Number(
                      detail.Base_Price || 0
                    ),

                  options:
                    String(
                      detail.Options || ''
                    ),

                  optionsPrice:
                    Number(
                      detail.Options_Price || 0
                    ),

                  itemTotal:
                    Number(
                      detail.Item_Total || 0
                    ),

                  remark:
                    String(
                      detail.Remark || ''
                    )

                };

              }
            )

        };

      }
    );


  return {

    success:
      true,

    date:
      today,

    orders:
      result

  };

}

function formatTimeValue_(value) {

  if (
    Object.prototype.toString
      .call(value) ===
      '[object Date]' &&
    !isNaN(value.getTime())
  ) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'HH:mm:ss'
    );

  }


  return String(
    value || ''
  ).trim();

}

function updateOrderStatus(
  orderId,
  orderDate,
  newStatus
) {

  const allowedStatuses = [
    'NEW',
    'PREPARING',
    'READY',
    'COMPLETED',
    'CANCELLED'
  ];


  orderId =
    String(orderId || '')
      .trim();


  orderDate =
    String(orderDate || '')
      .trim();


  newStatus =
    String(newStatus || '')
      .trim()
      .toUpperCase();


  if (!orderId) {

    throw new Error(
      'ไม่พบเลขออเดอร์'
    );

  }


  if (!orderDate) {

    throw new Error(
      'ไม่พบวันที่ออเดอร์'
    );

  }


  if (
    allowedStatuses.indexOf(
      newStatus
    ) === -1
  ) {

    throw new Error(
      'สถานะไม่ถูกต้อง: ' +
      newStatus
    );

  }


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const orderSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.ORDERS
    );


  if (!orderSheet) {

    throw new Error(
      'ไม่พบ Sheet Orders'
    );

  }


  const lastRow =
    orderSheet.getLastRow();


  if (lastRow < 2) {

    throw new Error(
      'ยังไม่มีออเดอร์'
    );

  }


  const values =
    orderSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        11
      )
      .getValues();


  let foundRow = 0;

  let oldStatus = '';


  values.forEach(
    function(row, index) {

      const currentId =
        String(
          row[0] || ''
        ).trim();


      const rawDate =
        row[1];


      let currentDate = '';


      if (
        Object.prototype
          .toString
          .call(rawDate) ===
          '[object Date]' &&
        !isNaN(
          rawDate.getTime()
        )
      ) {

        currentDate =
          Utilities.formatDate(
            rawDate,
            Session.getScriptTimeZone(),
            'yyyy-MM-dd'
          );

      } else {

        currentDate =
          String(
            rawDate || ''
          ).trim();

      }


      if (
        currentId === orderId &&
        currentDate === orderDate
      ) {

        foundRow =
          index + 2;


        oldStatus =
          String(
            row[10] || ''
          )
          .trim()
          .toUpperCase();

      }

    }
  );


  if (!foundRow) {

    throw new Error(
      'ไม่พบออเดอร์ ' +
      orderDate +
      ' / ' +
      orderId
    );

  }


  if (
    oldStatus ===
    newStatus
  ) {

    return {

      success: true,

      orderId:
        orderId,

      orderDate:
        orderDate,

      oldStatus:
        oldStatus,

      status:
        newStatus

    };

  }


  if (
    oldStatus === 'COMPLETED' ||
    oldStatus === 'CANCELLED'
  ) {

    throw new Error(
      'ออเดอร์ ' +
      orderId +
      ' ปิดงานแล้ว ไม่สามารถเปลี่ยนสถานะได้'
    );

  }


  orderSheet
    .getRange(
      foundRow,
      11
    )
    .setValue(
      newStatus
    );

    let notification = {
    success: true,
    sent: false,
    reason: 'NOT_READY'
  };

  if (
    newStatus === 'READY' &&
    oldStatus !== 'READY'
  ) {

    try {

      /*
       * Orders ปัจจุบัน:
       * D = LINE User ID
       */
      const rowData =
        values[foundRow - 2];

      const userId =
        String(
          rowData &&
          rowData[3]
            ? rowData[3]
            : ''
        ).trim();

      notification =
        sendLineReadyNotification_(
          userId,
          orderId,
          orderDate
        );

    } catch (error) {

      console.error(
        'LINE READY NOTIFICATION ERROR:',
        error
      );

      notification = {
        success: false,
        sent: false,
        reason:
          error.message ||
          String(error)
      };
    }
  }

  return {

    success: true,

    orderId:
      orderId,

    orderDate:
      orderDate,

    oldStatus:
      oldStatus,

    status:
      newStatus,

    notification: notification

  };

}

function getPickupTimeText_(
  slotMap,
  slotId
) {

  const id =
    String(
      slotId || ''
    ).trim();


  if (!id) {

    return '';

  }


  const slot =
    slotMap[id];


  if (!slot) {

    return id;

  }


  return (
    slot.start +
    ' - ' +
    slot.end
  );

}


function updatePaymentStatus(
  orderId,
  orderDate,
  newPaymentStatus
) {

  const allowedStatuses = [
    'UNPAID',
    'PENDING',
    'PAID',
    'REFUNDED'
  ];

  orderId = String(orderId || '').trim();
  orderDate = String(orderDate || '').trim();
  newPaymentStatus =
    String(newPaymentStatus || '')
      .trim()
      .toUpperCase();

  if (!orderId) {
    throw new Error('ไม่พบเลขออเดอร์');
  }

  if (!orderDate) {
    throw new Error('ไม่พบวันที่ออเดอร์');
  }

  if (allowedStatuses.indexOf(newPaymentStatus) === -1) {
    throw new Error(
      'สถานะการชำระเงินไม่ถูกต้อง: ' +
      newPaymentStatus
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(CONFIG.SHEETS.ORDERS);

  if (!sheet) {
    throw new Error('ไม่พบ Sheet Orders');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) {
    throw new Error('Orders ยังไม่มีข้อมูล');
  }

  const headers =
    sheet
      .getRange(1, 1, 1, lastCol)
      .getValues()[0]
      .map(function(header) {
        return String(header || '').trim();
      });

  const orderIdCol =
    headers.indexOf('Order_ID') + 1;

  const dateCol =
    headers.indexOf('Date') + 1;

  const paymentStatusCol =
    headers.indexOf('Payment_Status') + 1;

  if (orderIdCol <= 0) {
    throw new Error('ไม่พบคอลัมน์ Order_ID');
  }

  if (dateCol <= 0) {
    throw new Error('ไม่พบคอลัมน์ Date');
  }

  if (paymentStatusCol <= 0) {
    throw new Error('ไม่พบคอลัมน์ Payment_Status');
  }

  const values =
    sheet
      .getRange(2, 1, lastRow - 1, lastCol)
      .getValues();

  let foundRow = 0;
  let oldPaymentStatus = '';

  values.forEach(function(row, index) {

    if (foundRow) return;

    const currentId =
      String(row[orderIdCol - 1] || '').trim();

    const rawDate =
      row[dateCol - 1];

    let currentDate = '';

    if (
      Object.prototype.toString.call(rawDate) ===
        '[object Date]' &&
      !isNaN(rawDate.getTime())
    ) {
      currentDate =
        Utilities.formatDate(
          rawDate,
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        );
    } else {
      currentDate =
        String(rawDate || '').trim();
    }

    if (
      currentId === orderId &&
      currentDate === orderDate
    ) {

      foundRow = index + 2;

      oldPaymentStatus =
        String(
          row[paymentStatusCol - 1] || ''
        )
        .trim()
        .toUpperCase();

    }

  });

  if (!foundRow) {
    throw new Error(
      'ไม่พบออเดอร์ ' +
      orderDate +
      ' / ' +
      orderId
    );
  }

  sheet
    .getRange(foundRow, paymentStatusCol)
    .setValue(newPaymentStatus);

  SpreadsheetApp.flush();

  return {
    success: true,
    orderId: orderId,
    orderDate: orderDate,
    oldPaymentStatus: oldPaymentStatus,
    paymentStatus: newPaymentStatus
  };

}


/**
 * ทดสอบโครงสร้าง Payment โดยไม่แก้ข้อมูลออเดอร์
 */
function testPaymentStatus() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(CONFIG.SHEETS.ORDERS);

  if (!sheet) {
    throw new Error('ไม่พบ Sheet Orders');
  }

  const lastCol = sheet.getLastColumn();

  const headers =
    sheet
      .getRange(1, 1, 1, lastCol)
      .getValues()[0]
      .map(function(header) {
        return String(header || '').trim();
      });

  const required = [
    'Order_ID',
    'Date',
    'Payment_Status'
  ];

  const missing =
    required.filter(function(name) {
      return headers.indexOf(name) === -1;
    });

  if (missing.length) {
    throw new Error(
      'ขาดคอลัมน์: ' +
      missing.join(', ')
    );
  }

  Logger.log('PAYMENT SYSTEM OK');

  Logger.log(
    'Order_ID column = ' +
    (headers.indexOf('Order_ID') + 1)
  );

  Logger.log(
    'Date column = ' +
    (headers.indexOf('Date') + 1)
  );

  Logger.log(
    'Payment_Status column = ' +
    (headers.indexOf('Payment_Status') + 1)
  );

  Logger.log(
    'Allowed statuses = UNPAID, PENDING, PAID, REFUNDED'
  );

}


function getPaymentSettings_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'PAYMENT_SETTINGS'
    );

  if (!sheet) {

    throw new Error(
      'ไม่พบ Sheet PAYMENT_SETTINGS'
    );

  }


  const values =
    sheet
      .getDataRange()
      .getValues();


  const settings = {};


  values.forEach(
    function(row, index) {

      // ข้ามหัวตาราง
      if (index === 0) {
        return;
      }


      const key =
        String(
          row[0] || ''
        ).trim();


      if (!key) {
        return;
      }


      settings[key] =
        String(
          row[1] || ''
        ).trim();

    }
  );


  return {

    cashActive:
      settings.Cash_Active
        .toUpperCase() === 'TRUE',


    promptPayActive:
      settings.PromptPay_Active
        .toUpperCase() === 'TRUE',


    promptPayName:
      settings.PromptPay_Name || '',


    promptPayNumber:
      settings.PromptPay_Number || '',


    promptPayQrUrl:
      settings.PromptPay_QR_URL || '',

    promptPayQrData:
      getPaymentQRData_(
        settings.PromptPay_QR_URL || ''
      ),

    governmentActive:
      settings.Government_Active
        .toUpperCase() === 'TRUE',


    governmentName:
      settings.Government_Name || '',


    governmentNote:
      settings.Government_Note || ''

  };

}

function getPaymentQRData_(url) {

  if (!url) {
    return '';
  }

  try {

    const match =
      String(url).match(/[-\w]{25,}/);

    if (!match) {
      return '';
    }

    const fileId =
      match[0];

    /*
     * เปิดไฟล์จาก Drive
     */
    const file =
      DriveApp.getFileById(fileId);

    /*
     * อ่านไฟล์เป็น Blob
     */
    const blob =
      file.getBlob();

    /*
     * แปลงเป็น Base64
     */
    const base64 =
      Utilities.base64Encode(
        blob.getBytes()
      );

    const contentType =
      blob.getContentType() ||
      'image/png';

    return (
      'data:' +
      contentType +
      ';base64,' +
      base64
    );

  } catch (error) {

    console.log(
      'QR ERROR: ' +
      error.message
    );

    return '';

  }

}

function testPaymentQR() {

  const url =
    'https://drive.google.com/uc?export=view&id=1CkWayk0FWqJmg6j28IOWnVsSlGM9VuE3';

  const data =
    getPaymentQRData_(url);

  Logger.log(
    'QR DATA LENGTH = ' +
    (data ? data.length : 0)
  );

  if (data) {

    Logger.log(
      'SUCCESS: สร้าง QR Base64 สำเร็จ'
    );

    Logger.log(
      data.substring(0, 50)
    );

  } else {

    Logger.log(
      'ERROR: QR DATA ว่าง'
    );

  }

}

/* =========================================================
 * ADMIN SETTINGS - MENU + PAYMENT QR
 *
 * Uses existing CONFIG and PAYMENT_SETTINGS.
 * No CONFIG redeclaration.
 * ========================================================= */

function getAdminSettings() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const menuSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.MENU
    );

  if (!menuSheet) {
    throw new Error('ไม่พบ Sheet Menu');
  }

  const menus =
    getSheetObjects_(
      menuSheet
    );

  const paymentSheet =
    ss.getSheetByName(
      'PAYMENT_SETTINGS'
    );

  if (!paymentSheet) {
    throw new Error(
      'ไม่พบ Sheet PAYMENT_SETTINGS'
    );
  }

  const rows =
    paymentSheet
      .getDataRange()
      .getValues();

  const settings = {};

  rows.forEach(
    function(row, index) {

      if (index === 0) return;

      const key =
        String(
          row[0] || ''
        ).trim();

      if (!key) return;

      settings[key] =
        String(
          row[1] || ''
        ).trim();

    }
  );

  const qrUrl =
    settings.PromptPay_QR_URL || '';

  return {

    success: true,

    menus:
      menus.map(
        function(menu) {

          return {

            id:
              String(
                menu.Menu_ID || ''
              ),

            name:
              String(
                menu.Menu_Name || ''
              ),

            description:
              String(
                menu.Description || ''
              ),

            price:
              Number(
                menu.Price || 0
              ),

            category:
              String(
                menu.Category || ''
              ),

            image:
              String(
                menu.Image_URL || ''
              ),

            available:
              String(
                menu.Available || ''
              )
              .toUpperCase() ===
              'TRUE'

          };

        }
      ),

    payment: {

      cashActive:
        String(
          settings.Cash_Active ||
          'FALSE'
        )
        .toUpperCase() ===
        'TRUE',

      promptPayActive:
        String(
          settings.PromptPay_Active ||
          'FALSE'
        )
        .toUpperCase() ===
        'TRUE',

      promptPayName:
        settings.PromptPay_Name || '',

      promptPayNumber:
        settings.PromptPay_Number || '',

      promptPayQrUrl:
        qrUrl,

      promptPayQrData:
        qrUrl
          ? getPaymentQRData_(qrUrl)
          : '',

      governmentActive:
        String(
          settings.Government_Active ||
          'FALSE'
        )
        .toUpperCase() ===
        'TRUE',

      governmentName:
        settings.Government_Name || '',

      governmentNote:
        settings.Government_Note || ''

    }

  };

}


function findMenuHeaderColumn_(
  headers,
  name
) {

  const index =
    headers.indexOf(
      name
    );

  if (index < 0) {

    throw new Error(
      'ไม่พบคอลัมน์ ' +
      name +
      ' ใน Sheet Menu'
    );

  }

  return index + 1;

}


function saveMenu(
  menu
) {

  if (!menu) {
    throw new Error(
      'ไม่พบข้อมูลเมนู'
    );
  }

  const menuId =
    String(
      menu.id || ''
    ).trim();

  if (!menuId) {
    throw new Error(
      'ไม่พบ Menu_ID'
    );
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.MENU
      );

  if (!sheet) {
    throw new Error(
      'ไม่พบ Sheet Menu'
    );
  }

  const headers =
    getHeaders_(
      sheet
    );

  const lastRow =
    sheet.getLastRow();

  const lastCol =
    sheet.getLastColumn();

  const idCol =
    findMenuHeaderColumn_(
      headers,
      'Menu_ID'
    );

  if (lastRow < 2) {
    throw new Error(
      'Sheet Menu ยังไม่มีข้อมูล'
    );
  }

  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastCol
      )
      .getValues();

  let targetRow = 0;

  rows.forEach(
    function(row, index) {

      if (targetRow) return;

      if (
        String(
          row[idCol - 1] || ''
        ).trim() ===
        menuId
      ) {
        targetRow =
          index + 2;
      }

    }
  );

  if (!targetRow) {
    throw new Error(
      'ไม่พบเมนู ' +
      menuId
    );
  }

  const values = {

    Menu_Name:
      String(
        menu.name || ''
      ).trim(),

    Description:
      String(
        menu.description || ''
      ).trim(),

    Price:
      Number(
        menu.price || 0
      ),

    Category:
      String(
        menu.category || ''
      ).trim(),

    Image_URL:
      String(
        menu.image || ''
      ).trim(),

    Available:
      Boolean(
        menu.available
      )

  };

  Object.keys(
    values
  ).forEach(
    function(key) {

      const col =
        headers.indexOf(
          key
        ) + 1;

      if (col > 0) {

        sheet
          .getRange(
            targetRow,
            col
          )
          .setValue(
            values[key]
          );

      }

    }
  );

  SpreadsheetApp.flush();

  return {
    success: true,
    menuId: menuId
  };

}


function toggleMenuAvailable(
  menuId,
  available
) {

  const id =
    String(
      menuId || ''
    ).trim();

  if (!id) {
    throw new Error(
      'ไม่พบ Menu_ID'
    );
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.MENU
      );

  if (!sheet) {
    throw new Error(
      'ไม่พบ Sheet Menu'
    );
  }

  const headers =
    getHeaders_(
      sheet
    );

  const idCol =
    findMenuHeaderColumn_(
      headers,
      'Menu_ID'
    );

  const availableCol =
    findMenuHeaderColumn_(
      headers,
      'Available'
    );

  const lastRow =
    sheet.getLastRow();

  const lastCol =
    sheet.getLastColumn();

  if (lastRow < 2) {
    throw new Error(
      'Sheet Menu ยังไม่มีข้อมูล'
    );
  }

  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastCol
      )
      .getValues();

  let targetRow = 0;

  rows.forEach(
    function(row, index) {

      if (targetRow) return;

      if (
        String(
          row[idCol - 1] || ''
        ).trim() ===
        id
      ) {
        targetRow =
          index + 2;
      }

    }
  );

  if (!targetRow) {
    throw new Error(
      'ไม่พบเมนู ' +
      id
    );
  }

  sheet
    .getRange(
      targetRow,
      availableCol
    )
    .setValue(
      Boolean(
        available
      )
    );

  SpreadsheetApp.flush();

  return {
    success: true,
    menuId: id,
    available: Boolean(available)
  };

}


function createMenu(
  menu
) {

  if (!menu) {
    throw new Error(
      'ไม่พบข้อมูลเมนู'
    );
  }

  const name =
    String(
      menu.name || ''
    ).trim();

  if (!name) {
    throw new Error(
      'กรุณาใส่ชื่อเมนู'
    );
  }

  const price =
    Number(
      menu.price || 0
    );

  if (
    !isFinite(price) ||
    price < 0
  ) {
    throw new Error(
      'ราคาไม่ถูกต้อง'
    );
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.MENU
      );

  if (!sheet) {
    throw new Error(
      'ไม่พบ Sheet Menu'
    );
  }

  const headers =
    getHeaders_(
      sheet
    );

  const lastRow =
    sheet.getLastRow();

  const lastCol =
    sheet.getLastColumn();

  const idCol =
    findMenuHeaderColumn_(
      headers,
      'Menu_ID'
    );

  const ids =
    lastRow >= 2
      ? sheet
          .getRange(
            2,
            idCol,
            lastRow - 1,
            1
          )
          .getValues()
      : [];

  let maxNumber = 0;

  ids.forEach(
    function(row) {

      const match =
        String(
          row[0] || ''
        )
        .match(
          /(\d+)$/
        );

      if (match) {

        maxNumber =
          Math.max(
            maxNumber,
            Number(
              match[1]
            )
          );

      }

    }
  );

  const newId =
    'M' +
    String(
      maxNumber + 1
    )
    .padStart(
      3,
      '0'
    );

  const row =
    Array(
      lastCol
    ).fill('');

  const values = {

    Menu_ID:
      newId,

    Menu_Name:
      name,

    Description:
      String(
        menu.description || ''
      ).trim(),

    Price:
      price,

    Category:
      String(
        menu.category || ''
      ).trim(),

    Image_URL:
      String(
        menu.image || ''
      ).trim(),

    Available:
      true

  };

  Object.keys(
    values
  ).forEach(
    function(key) {

      const index =
        headers.indexOf(
          key
        );

      if (index >= 0) {
        row[index] =
          values[key];
      }

    }
  );

  sheet.appendRow(
    row
  );

  SpreadsheetApp.flush();

  return {
    success: true,
    menu: {
      id: newId,
      name: name,
      price: price,
      description:
        values.Description,
      category:
        values.Category,
      image:
        values.Image_URL,
      available: true
    }
  };

}


function savePaymentSettings(
  settings
) {

  if (!settings) {
    throw new Error(
      'ไม่พบข้อมูลการชำระเงิน'
    );
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        'PAYMENT_SETTINGS'
      );

  if (!sheet) {
    throw new Error(
      'ไม่พบ Sheet PAYMENT_SETTINGS'
    );
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error(
      'PAYMENT_SETTINGS ยังไม่มีข้อมูล'
    );
  }

  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getValues();

  const incoming = {

    Cash_Active:
      settings.cashActive
        ? 'TRUE'
        : 'FALSE',

    PromptPay_Active:
      settings.promptPayActive
        ? 'TRUE'
        : 'FALSE',

    PromptPay_Name:
      String(
        settings.promptPayName || ''
      ).trim(),

    PromptPay_Number:
      String(
        settings.promptPayNumber || ''
      ).trim(),

    PromptPay_QR_URL:
      String(
        settings.promptPayQrUrl || ''
      ).trim(),

    Government_Active:
      settings.governmentActive
        ? 'TRUE'
        : 'FALSE',

    Government_Name:
      String(
        settings.governmentName || ''
      ).trim(),

    Government_Note:
      String(
        settings.governmentNote || ''
      ).trim()

  };

  rows.forEach(
    function(row, index) {

      const key =
        String(
          row[0] || ''
        ).trim();

      if (
        Object.prototype.hasOwnProperty
          .call(
            incoming,
            key
          )
      ) {

        sheet
          .getRange(
            index + 2,
            2
          )
          .setValue(
            incoming[key]
          );

      }

    }
  );

  SpreadsheetApp.flush();

  return {
    success: true,
    settings: incoming
  };

}


function uploadAdminImage(image) {

  if (!image) {
    throw new Error('ไม่พบรูปภาพ');
  }

  const dataUrl =
    String(image.dataUrl || '').trim();

  if (!dataUrl) {
    throw new Error('ไม่พบข้อมูลรูปภาพ');
  }

  const match =
    dataUrl.match(
      /^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/
    );

  if (!match) {
    throw new Error(
      'รูปภาพต้องเป็น Data URL'
    );
  }

  const mimeType =
    String(match[1]).toLowerCase();

  const supportedTypes = [
    'image/jpeg',
    'image/png',
    'image/webp'
  ];

  if (
    supportedTypes.indexOf(mimeType) === -1
  ) {
    throw new Error(
      'รองรับเฉพาะ JPG, PNG และ WebP'
    );
  }

  const bytes =
    Utilities.base64Decode(match[2]);

  if (!bytes || !bytes.length) {
    throw new Error(
      'ไม่สามารถอ่านข้อมูลรูปภาพได้'
    );
  }

  if (
    bytes.length >
    4 * 1024 * 1024
  ) {
    throw new Error(
      'รูปใหญ่เกิน 4 MB'
    );
  }

  const extension =
    githubImageExtensionProd_(
      mimeType
    );

  const originalName =
    String(
      image.fileName ||
      'menu-image'
    )
    .replace(
      /[^a-zA-Z0-9._-]/g,
      '_'
    );

  const baseName =
    originalName
      .replace(
        /\.[^.]+$/,
        ''
      )
      .replace(
        /[^a-zA-Z0-9_-]/g,
        '_'
      )
      .substring(
        0,
        60
      ) ||
    'menu-image';

  const uniqueName =
    baseName +
    '-' +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyyMMdd-HHmmss'
    ) +
    '-' +
    Utilities.getUuid()
      .replace(/-/g, '')
      .substring(0, 8) +
    '.' +
    extension;

  const imageDir =
    GITHUB_PROD_IMAGE_DIR_.replace(/\/+$/, '');

  const repoPath =
    imageDir + '/' +
    uniqueName;

  const result =
    uploadFileToGitHubProd_(
      repoPath,
      bytes,
      mimeType
    );

  return {
    success: true,

    fileId:
      result.sha || '',

    url:
      'https://raw.githubusercontent.com/' +
      GITHUB_PROD_OWNER_ +
      '/' +
      GITHUB_PROD_REPO_ +
      '/' +
      GITHUB_PROD_BRANCH_ +
      '/' +
      repoPath,

    name:
      uniqueName,

    path:
      repoPath,

    storage:
      'GITHUB'
  };
}

const GITHUB_PROD_OWNER_ =
  getStoreProperty_('GITHUB_OWNER', '');

const GITHUB_PROD_REPO_ =
  getStoreProperty_('GITHUB_REPO', '');

const GITHUB_PROD_BRANCH_ =
  getStoreProperty_('GITHUB_BRANCH', 'main');

const GITHUB_PROD_IMAGE_DIR_ =
  getStoreProperty_('GITHUB_IMAGE_DIR', 'images/menu');


function githubImageExtensionProd_(
  mimeType
) {

  const map = {
    'image/jpeg': 'jpg',
    'image/png': 'png',
    'image/webp': 'webp'
  };

  return (
    map[mimeType] ||
    'jpg'
  );
}


function uploadFileToGitHubProd_(
  repoPath,
  bytes,
  mimeType
) {

  const props =
    PropertiesService
      .getScriptProperties();

  const token =
    String(
      props.getProperty(
        'GITHUB_TOKEN'
      ) || ''
    ).trim();

  if (!token) {
    throw new Error(
      'ไม่พบ GITHUB_TOKEN ใน Script Properties ของ Store Instance'
    );
  }

  const apiUrl =
    'https://api.github.com/repos/' +
    encodeURIComponent(
      GITHUB_PROD_OWNER_
    ) +
    '/' +
    encodeURIComponent(
      GITHUB_PROD_REPO_
    ) +
    '/contents/' +
    repoPath
      .split('/')
      .map(function(part) {
        return encodeURIComponent(part);
      })
      .join('/');

  const payload = {
    message:
      'Upload menu image ' +
      repoPath,

    content:
      Utilities.base64Encode(
        bytes
      ),

    branch:
      GITHUB_PROD_BRANCH_
  };

  const response =
    UrlFetchApp.fetch(
      apiUrl,
      {
        method: 'put',

        contentType:
          'application/json',

        headers: {
          Authorization:
            'Bearer ' + token,

          Accept:
            'application/vnd.github+json',

          'X-GitHub-Api-Version':
            '2026-03-10'
        },

        payload:
          JSON.stringify(
            payload
          ),

        muteHttpExceptions:
          true
      }
    );

  const status =
    response.getResponseCode();

  const text =
    response.getContentText();

  let result;

  try {

    result =
      JSON.parse(text);

  } catch (error) {

    throw new Error(
      'GitHub ตอบกลับไม่ใช่ JSON (' +
      status +
      ')'
    );

  }

  if (
    status < 200 ||
    status >= 300
  ) {

    const detail =
      result &&
      result.message
        ? result.message
        : text;

    throw new Error(
      'อัปโหลดรูปไป GitHub ไม่สำเร็จ (' +
      status +
      '): ' +
      detail
    );
  }

  return {
    success: true,

    sha:
      result &&
      result.content &&
      result.content.sha
        ? result.content.sha
        : '',

    htmlUrl:
      result &&
      result.content &&
      result.content.html_url
        ? result.content.html_url
        : '',

    downloadUrl:
      result &&
      result.content &&
      result.content.download_url
        ? result.content.download_url
        : ''
  };
}

/* =========================================================
 * ADMIN BACKEND AUTHENTICATION
 *
 * PIN is stored in Script Properties.
 * Session tokens are also stored in Script Properties
 * with an expiry timestamp.
 * Default PIN is initialized by setupAdminSecurity()
 * and should be changed before production if desired.
 * ========================================================= */

const ADMIN_AUTH_PIN_KEY_ =
  'ADMIN_AUTH_PIN';

const ADMIN_AUTH_SESSIONS_KEY_ =
  'ADMIN_AUTH_SESSIONS';


function setupAdminSecurity() {

  const props =
    PropertiesService
      .getScriptProperties();

  const existing =
    props.getProperty(
      ADMIN_AUTH_PIN_KEY_
    );


  if (!existing) {

    props.setProperty(
      ADMIN_AUTH_PIN_KEY_,
      '8899'
    );

    return {

      success: true,

      message:
        'ตั้งค่า Admin PIN 8899 สำเร็จ'

    };

  }


  return {

    success: true,

    message:
      'มี Admin PIN อยู่แล้ว'

  };

}


function loginAdmin(
  pin
) {

  const input =
    String(
      pin || ''
    ).trim();


  const props =
    PropertiesService
      .getScriptProperties();


  const savedPin =
    props.getProperty(
      ADMIN_AUTH_PIN_KEY_
    );


  if (!savedPin) {

    throw new Error(
      'ยังไม่ได้ตั้งค่า Admin PIN กรุณารัน setupAdminSecurity() 1 ครั้ง'
    );

  }


  if (
    input !==
    savedPin
  ) {

    throw new Error(
      'รหัส Admin ไม่ถูกต้อง'
    );

  }


  const token =
    Utilities.getUuid();


  const expiresAt =
    Date.now() +
    8 * 60 * 60 * 1000;


  const raw =
    props.getProperty(
      ADMIN_AUTH_SESSIONS_KEY_
    );


  let sessions = {};

  try {

    sessions =
      raw
        ? JSON.parse(raw)
        : {};

  } catch (error) {

    sessions = {};

  }


  // Clean expired sessions.
  Object.keys(
    sessions
  ).forEach(
    function(key) {

      if (
        !sessions[key] ||
        Number(
          sessions[key]
        ) <=
        Date.now()
      ) {

        delete sessions[key];

      }

    }
  );


  sessions[token] =
    expiresAt;


  props.setProperty(
    ADMIN_AUTH_SESSIONS_KEY_,
    JSON.stringify(
      sessions
    )
  );


  return {

    success: true,

    token:
      token,

    expiresAt:
      expiresAt

  };

}


function requireAdminAuth_(
  token
) {

  const authToken =
    String(
      token || ''
    ).trim();


  if (!authToken) {

    throw new Error(
      'ต้องเข้าสู่ระบบ Admin ก่อน'
    );

  }


  const props =
    PropertiesService
      .getScriptProperties();


  const raw =
    props.getProperty(
      ADMIN_AUTH_SESSIONS_KEY_
    );


  if (!raw) {

    throw new Error(
      'Admin Session หมดอายุ กรุณาเข้าสู่ระบบใหม่'
    );

  }


  let sessions = {};

  try {

    sessions =
      JSON.parse(
        raw
      );

  } catch (error) {

    throw new Error(
      'Admin Session ไม่ถูกต้อง'
    );

  }


  const expiresAt =
    Number(
      sessions[authToken] || 0
    );


  if (
    !expiresAt ||
    expiresAt <=
    Date.now()
  ) {

    delete sessions[
      authToken
    ];


    props.setProperty(
      ADMIN_AUTH_SESSIONS_KEY_,
      JSON.stringify(
        sessions
      )
    );


    throw new Error(
      'Admin Session หมดอายุ กรุณาเข้าสู่ระบบใหม่'
    );

  }


  return true;

}


function logoutAdmin(
  token
) {

  const authToken =
    String(
      token || ''
    ).trim();


  const props =
    PropertiesService
      .getScriptProperties();


  const raw =
    props.getProperty(
      ADMIN_AUTH_SESSIONS_KEY_
    );


  if (!raw) {

    return {
      success: true
    };

  }


  let sessions = {};

  try {

    sessions =
      JSON.parse(
        raw
      );

  } catch (error) {

    sessions = {};

  }


  delete sessions[
    authToken
  ];


  props.setProperty(
    ADMIN_AUTH_SESSIONS_KEY_,
    JSON.stringify(
      sessions
    )
  );


  return {
    success: true
  };

}


function changeAdminPin(
  currentPin,
  newPin,
  token
) {

  requireAdminAuth_(
    token
  );


  const current =
    String(
      currentPin || ''
    ).trim();


  const next =
    String(
      newPin || ''
    ).trim();


  const props =
    PropertiesService
      .getScriptProperties();


  const saved =
    props.getProperty(
      ADMIN_AUTH_PIN_KEY_
    );

  if (!saved) {

    throw new Error(
      'ยังไม่ได้ตั้งค่า Admin PIN กรุณารัน setupAdminSecurity() ก่อน'
    );

  }


  if (
    current !==
    saved
  ) {

    throw new Error(
      'รหัสเดิมไม่ถูกต้อง'
    );

  }


  if (
    !/^\d{4,6}$/.test(
      next
    )
  ) {

    throw new Error(
      'PIN ใหม่ต้องเป็นตัวเลข 4-6 หลัก'
    );

  }


  props.setProperty(
    ADMIN_AUTH_PIN_KEY_,
    next
  );


  return {
    success: true
  };

}



/* =========================================================
 * MENU SYSTEM V2 HELPERS
 * ========================================================= */

function getMenuSystemData() {
  const data = getInitialData();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const menuSheet = ss.getSheetByName(CONFIG.SHEETS.MENU);
  const optionSheet = ss.getSheetByName(CONFIG.SHEETS.OPTIONS);
  const dailySheet = ss.getSheetByName('Daily_Dishes');
  const specialSheet = ss.getSheetByName('Special_Menus');

  const menus = menuSheet ? getSheetObjects_(menuSheet).map(function(menu) {
    return {
      id: String(menu.Menu_ID || ''),
      name: String(menu.Menu_Name || ''),
      description: String(menu.Description || ''),
      price: Number(menu.Price || 0),
      category: String(menu.Category || ''),
      image: String(menu.Image_URL || ''),
      available: String(menu.Available || '').toUpperCase() === 'TRUE'
    };
  }) : [];

  const options = optionSheet ? getSheetObjects_(optionSheet).map(function(option) {
    return {
      menuId: String(option.Menu_ID || ''),
      group: String(option.Option_Group || ''),
      name: String(option.Option_Name || ''),
      price: Number(option.Option_Price || 0),
      available: String(option.Available || '').toUpperCase() === 'TRUE',
      minSelect: Number(option.Min_Select || 0),
      maxSelect: Number(option.Max_Select || 0)
    };
  }) : [];

  const dailyDishes = dailySheet
  ? getSheetObjects_(dailySheet)
      .map(function(dish) {
        return {
          id: String(dish.Dish_ID || ''),
          date: normalizeDateValue_(dish.Dish_Date),
          name: String(dish.Dish_Name || ''),
          description: String(dish.Description || ''),
          image: String(dish.Image_URL || ''),
          available:
            String(dish.Available || '').toUpperCase() === 'TRUE'
        };
      })
      .filter(function(dish) {
        return dish.date && dish.name;
      })
  : [];

  const specialMenus = specialSheet ? getSheetObjects_(specialSheet).map(function(menu) {
    return {
      id: String(menu.Special_ID || ''),
      startDate: String(menu.Start_Date || ''),
      endDate: String(menu.End_Date || ''),
      name: String(menu.Menu_Name || ''),
      description: String(menu.Description || ''),
      price: Number(menu.Price || 0),
      image: String(menu.Image_URL || ''),
      available: String(menu.Available || '').toUpperCase() === 'TRUE'
    };
  }) : [];

  return {
    success: true,
    today: today,
    menus: menus,
    options: options,
    dailyDishes: dailyDishes,
    specialMenus: specialMenus
  };
}

function ensureSheetWithHeaders_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function backupSheetIfExists_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return '';
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  let name = sheetName + '_BACKUP_' + stamp;
  let i = 1;
  while (ss.getSheetByName(name)) {
    name = sheetName + '_BACKUP_' + stamp + '_' + i++;
  }
  sheet.copyTo(ss).setName(name);
  return name;
}

function installRecommendedMenuSystem(confirm) {
  if (String(confirm).toUpperCase() !== 'YES') {
    throw new Error('ต้องส่ง confirm = YES เพื่อยืนยันการติดตั้งชุดเมนูใหม่');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupMenu = backupSheetIfExists_(CONFIG.SHEETS.MENU);
  const backupOptions = backupSheetIfExists_(CONFIG.SHEETS.OPTIONS);

  const menuSheet = ss.getSheetByName(CONFIG.SHEETS.MENU);
  const optionSheet = ss.getSheetByName(CONFIG.SHEETS.OPTIONS);

  const menuHeaders = ['Menu_ID','Menu_Name','Description','Price','Category','Image_URL','Available'];
  const optionHeaders = ['Menu_ID','Option_Group','Option_Name','Option_Price','Available','Min_Select','Max_Select'];

  menuSheet.clearContents();
  optionSheet.clearContents();
  menuSheet.getRange(1,1,1,menuHeaders.length).setValues([menuHeaders]);
  optionSheet.getRange(1,1,1,optionHeaders.length).setValues([optionHeaders]);

  const mainMenus = [
    ['M001','กะเพราไก่','เผ็ดปกติ',50,'อาหารตามสั่ง','',true],
    ['M002','กะเพราหมู','',50,'อาหารตามสั่ง','',true],
    ['M003','กะเพราเนื้อ','',60,'อาหารตามสั่ง','',true],
    ['M004','กะเพราทะเล','',70,'อาหารตามสั่ง','',true],
    ['M005','ผัดพริกแกงหมู','',50,'อาหารตามสั่ง','',true],
    ['M006','ผัดพริกแกงไก่','',50,'อาหารตามสั่ง','',true],
    ['M007','คะน้าหมูกรอบ','',60,'อาหารตามสั่ง','',true],
    ['M008','ผัดผักรวมหมู','',50,'อาหารตามสั่ง','',true],
    ['M009','หมูกระเทียม','',50,'อาหารตามสั่ง','',true],
    ['M010','ไก่กระเทียม','',50,'อาหารตามสั่ง','',true],
    ['M011','ผัดพริกสดหมู','',50,'อาหารตามสั่ง','',true],
    ['M012','ผัดขี้เมาทะเล','',70,'อาหารตามสั่ง','',true],
    ['M013','ข้าวผัดหมู','',50,'อาหารตามสั่ง','',true],
    ['M014','ข้าวผัดไก่','',50,'อาหารตามสั่ง','',true],
    ['M015','ข้าวผัดกุ้ง','',60,'อาหารตามสั่ง','',true],
    ['M016','ผัดซีอิ๊วหมู','',50,'อาหารตามสั่ง','',true],
    ['M017','ราดหน้าหมู','',50,'อาหารตามสั่ง','',true],
    ['M018','ต้มยำน้ำข้น','',60,'อาหารตามสั่ง','',true],
    ['M019','ต้มยำน้ำใส','',60,'อาหารตามสั่ง','',true],
    ['M020','ต้มจืดเต้าหู้หมูสับ','',50,'อาหารตามสั่ง','',true],
    ['M021','ไข่เจียวหมูสับราดข้าว','',50,'อาหารตามสั่ง','',true],
    ['M022','ข้าวราดแกง 1 อย่าง','เลือกกับข้าว 1 อย่าง',35,'ข้าวราดแกง','',true],
    ['M023','ข้าวราดแกง 2 อย่าง','เลือกกับข้าว 2 อย่าง',40,'ข้าวราดแกง','',true],
    ['M024','ข้าวราดแกง 3 อย่าง','เลือกกับข้าว 3 อย่าง',45,'ข้าวราดแกง','',true]
  ];
  menuSheet.getRange(2,1,mainMenus.length,menuHeaders.length).setValues(mainMenus);

  const optionRows = [];
  const spicyMenus = ['M001','M002','M003','M004','M005','M006','M011','M012','M018','M019'];
  spicyMenus.forEach(function(menuId) {
    [['ไม่เผ็ด',0],['เผ็ดน้อย',0],['เผ็ดปกติ',0],['เผ็ดมาก',0]].forEach(function(x) {
      optionRows.push([menuId,'ระดับความเผ็ด',x[0],x[1],true,0,1]);
    });
  });

  mainMenus.forEach(function(row) {
    const id=row[0];
    if (['M022','M023','M024'].indexOf(id)>=0) return;
    optionRows.push([id,'เพิ่ม','ไข่ดาว',10,true,0,1]);
    optionRows.push([id,'เพิ่ม','ไข่เจียว',15,true,0,1]);
    optionRows.push([id,'เพิ่ม','เพิ่มข้าว',10,true,0,1]);
  });
  optionSheet.getRange(2,1,optionRows.length,optionHeaders.length).setValues(optionRows);

  ensureSheetWithHeaders_('Daily_Dishes', ['Dish_ID','Dish_Date','Dish_Name','Description','Image_URL','Available']);
  ensureSheetWithHeaders_('Special_Menus', ['Special_ID','Start_Date','End_Date','Menu_Name','Description','Price','Image_URL','Available']);
  SpreadsheetApp.flush();

  return {success:true, mainMenus:mainMenus.length, options:optionRows.length, backupMenu:backupMenu, backupOptions:backupOptions};
}

function getDailyDishLibrary() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.SHEETS.DAILY_DISH_LIBRARY
    );

  if (!sheet || sheet.getLastRow() < 2) {

    return {
      success: true,
      menus: []
    };

  }

  const rows =
    getSheetObjects_(sheet);

  return {

    success: true,

    menus:
      rows
        .map(function(row) {

          return {

            libraryId:
              String(
                row.Library_ID || ''
              ),

            name:
              String(
                row.Dish_Name || ''
              ),

            description:
              String(
                row.Description || ''
              )

          };

        })
        .filter(function(row) {

          return row.name;

        })

  };
}


function addDailyDishesFromLibrary(
  dateValue,
  libraryIds
) {

  const date =
    String(
      dateValue || ''
    ).trim();


  if (!date) {

    throw new Error(
      'กรุณาระบุวันที่'
    );

  }


  const ids =
    Array.isArray(
      libraryIds
    )
      ? libraryIds.map(
          String
        )
      : [];


  if (!ids.length) {

    throw new Error(
      'กรุณาเลือกเมนูอย่างน้อย 1 รายการ'
    );

  }


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const librarySheet =
  ss.getSheetByName(
    CONFIG.SHEETS.DAILY_DISH_LIBRARY
  );


  if (!librarySheet) {

    throw new Error(
      'ไม่พบ Sheet Daily_Dish_Library'
    );

  }


  const dailySheet =
    ensureSheetWithHeaders_(
      'Daily_Dishes',
      [
        'Dish_ID',
        'Dish_Date',
        'Dish_Name',
        'Description',
        'Image_URL',
        'Available'
      ]
    );


  const library =
    getSheetObjects_(
      librarySheet
    );


  const selected =
    library.filter(
      function(row) {

        return ids.indexOf(
          String(
            row.Library_ID || ''
          )
        ) >= 0;

      }
    );


  if (!selected.length) {

    throw new Error(
      'ไม่พบเมนูจากคลังที่เลือก'
    );

  }


  const existing =
    dailySheet.getLastRow() >= 2
      ? getSheetObjects_(
          dailySheet
        )
      : [];


  const existingKeys = {};


  existing.forEach(
    function(row) {

      const key =
        normalizeDateValue_(row.Dish_Date) +
        '|' +
        String(
          row.Dish_Name || ''
        ).trim();


      existingKeys[key] =
        true;

    }
  );


  let maxId = 0;


  existing.forEach(
    function(row) {

      const match =
        String(
          row.Dish_ID || ''
        ).match(
          /(\d+)$/
        );


      if (match) {

        maxId =
          Math.max(
            maxId,
            Number(
              match[1]
            )
          );

      }

    }
  );


  const rows = [];

  const added = [];

  const skipped = [];


  selected.forEach(
    function(row) {

      const name =
        String(
          row.Dish_Name || ''
        ).trim();


      const key =
        normalizeDateValue_(date) +
        '|' +
        name;


      // มีเมนูนี้ในวันเดียวกันแล้ว
      if (
        existingKeys[key]
      ) {

        skipped.push(
          name
        );

        return;

      }


      maxId += 1;


      rows.push([
        'D' + String(maxId).padStart(3, '0'),
        date,
        name,
        String(row.Description || ''),
        String(row.Image_URL || row.Image || ''),
        true
      ]);


            existingKeys[key] =
              true;


            added.push(
              name
            );

          }
        );


  if (rows.length) {

    dailySheet
      .getRange(
        dailySheet.getLastRow() + 1,
        1,
        rows.length,
        6
      )
      .setValues(
        rows
      );

  }


  SpreadsheetApp.flush();


  return {

    success: true,

    date:
      date,

    added:
      added,

    skipped:
      skipped,

    addedCount:
      added.length,

    skippedCount:
      skipped.length

  };

}

function createDailyDish(dish) {
  if (!dish) throw new Error('ไม่พบข้อมูลกับข้าว');
  const name=String(dish.name||'').trim();
  const date=String(dish.date||'').trim();
  if (!name) throw new Error('กรุณาใส่ชื่อกับข้าว');
  if (!date) throw new Error('กรุณาเลือกวันที่');
  const sheet=ensureSheetWithHeaders_('Daily_Dishes',['Dish_ID','Dish_Date','Dish_Name','Description','Image_URL','Available']);
  const rows=sheet.getLastRow()>=2?sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues():[];
  let max=0;
  rows.forEach(function(r){const m=String(r[0]||'').match(/(\d+)$/); if(m) max=Math.max(max,Number(m[1]));});
  const id='D'+String(max+1).padStart(3,'0');
  sheet.appendRow([id,date,name,String(dish.description||'').trim(),String(dish.image||'').trim(),true]);
  SpreadsheetApp.flush();
  return {success:true,id:id};
}

function saveDailyDish(dish) {
  if (!dish) throw new Error('ไม่พบข้อมูลกับข้าว');
  return updateSimpleRow_('Daily_Dishes','Dish_ID',dish.id,{
    Dish_Date:String(dish.date||'').trim(),
    Dish_Name:String(dish.name||'').trim(),
    Description:String(dish.description||'').trim(),
    Image_URL:String(dish.image||'').trim(),
    Available:Boolean(dish.available)
  });
}

function toggleDailyDish(dishId,available){return setAvailabilityById_('Daily_Dishes','Dish_ID',dishId,'Available',available);}
function deleteDailyDish(dishId){return deleteById_('Daily_Dishes','Dish_ID',dishId);}

function createSpecialMenu(menu) {
  if (!menu) throw new Error('ไม่พบข้อมูลเมนูพิเศษ');
  const name=String(menu.name||'').trim();
  if (!name) throw new Error('กรุณาใส่ชื่อเมนูพิเศษ');
  const sheet=ensureSheetWithHeaders_('Special_Menus',['Special_ID','Start_Date','End_Date','Menu_Name','Description','Price','Image_URL','Available']);
  const rows=sheet.getLastRow()>=2?sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues():[];
  let max=0;
  rows.forEach(function(r){const m=String(r[0]||'').match(/(\d+)$/); if(m) max=Math.max(max,Number(m[1]));});
  const id='S'+String(max+1).padStart(3,'0');
  const start=String(menu.startDate||'').trim();
  const end=String(menu.endDate||start).trim();
  sheet.appendRow([id,start,end,name,String(menu.description||'').trim(),Number(menu.price||0),String(menu.image||'').trim(),true]);
  SpreadsheetApp.flush();
  return {success:true,id:id};
}

function saveSpecialMenu(menu) {
  if (!menu) throw new Error('ไม่พบข้อมูลเมนูพิเศษ');
  return updateSimpleRow_('Special_Menus','Special_ID',menu.id,{
    Start_Date:String(menu.startDate||'').trim(),
    End_Date:String(menu.endDate||'').trim(),
    Menu_Name:String(menu.name||'').trim(),
    Description:String(menu.description||'').trim(),
    Price:Number(menu.price||0),
    Image_URL:String(menu.image||'').trim(),
    Available:Boolean(menu.available)
  });
}

function toggleSpecialMenu(menuId,available){return setAvailabilityById_('Special_Menus','Special_ID',menuId,'Available',available);}

function updateSimpleRow_(sheetName,idHeader,id,values) {
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('ไม่พบ Sheet '+sheetName);
  const headers=getHeaders_(sheet);
  const idCol=headers.indexOf(idHeader)+1;
  if (idCol<=0) throw new Error('ไม่พบคอลัมน์ '+idHeader);
  const lastRow=sheet.getLastRow(), lastCol=sheet.getLastColumn();
  if (lastRow<2) throw new Error('Sheet '+sheetName+' ยังไม่มีข้อมูล');
  const rows=sheet.getRange(2,1,lastRow-1,lastCol).getValues();
  let rowNumber=-1;
  rows.forEach(function(row,index){ if(String(row[idCol-1]||'').trim()===String(id||'').trim()) rowNumber=index+2; });
  if(rowNumber<0) throw new Error('ไม่พบข้อมูล '+id);
  Object.keys(values).forEach(function(key){const col=headers.indexOf(key)+1;if(col>0) sheet.getRange(rowNumber,col).setValue(values[key]);});
  SpreadsheetApp.flush();
  return {success:true,id:String(id)};
}

function setAvailabilityById_(sheetName,idHeader,id,availableHeader,available) {
  return updateSimpleRow_(sheetName,idHeader,id,{[availableHeader]:Boolean(available)});
}

function deleteById_(sheetName,idHeader,id) {
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(!sheet) throw new Error('ไม่พบ Sheet '+sheetName);
  const headers=getHeaders_(sheet); const idCol=headers.indexOf(idHeader)+1;
  const rows=sheet.getLastRow()>=2?sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues():[];
  let rowNumber=-1; rows.forEach(function(row,index){if(String(row[idCol-1]||'').trim()===String(id||'').trim()) rowNumber=index+2;});
  if(rowNumber<0) throw new Error('ไม่พบข้อมูล '+id);
  sheet.deleteRow(rowNumber); SpreadsheetApp.flush(); return {success:true,id:String(id)};
}

function testDailyDishLibrary() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.SHEETS.DAILY_DISH_LIBRARY
    );

  if (!sheet) {

    throw new Error(
      'ไม่พบ Sheet: ' +
      CONFIG.SHEETS.DAILY_DISH_LIBRARY
    );

  }

  const data =
    getSheetObjects_(sheet);

  Logger.log(
    'Sheet = ' +
    sheet.getName()
  );

  Logger.log(
    'จำนวนเมนู = ' +
    data.length
  );

  Logger.log(
    JSON.stringify(data)
  );

  return data;
}

function testGetDailyDishLibrary() {

  const result =
    getDailyDishLibrary();

  Logger.log(
    JSON.stringify(
      result
    )
  );

}

function cleanupDuplicateDailyDishes() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Daily_Dishes');

  if (!sheet) {
    throw new Error(
      'ไม่พบ Sheet Daily_Dishes'
    );
  }

  if (sheet.getLastRow() < 2) {
    return {
      success: true,
      removed: 0
    };
  }

  const data =
    getSheetObjects_(sheet);

  const seen = {};
  const rowsToDelete = [];

  data.forEach(function(row, index) {

    const date =
      normalizeDateValue_(
        row.Dish_Date
      );

    const name =
      String(
        row.Dish_Name || ''
      ).trim();

    if (!date || !name) {
      return;
    }

    const key =
      date + '|' + name;

    if (!seen[key]) {

      seen[key] = {
        sheetRow:
          index + 2,

        available:
          String(
            row.Available || ''
          ).toUpperCase() === 'TRUE'
      };

      return;
    }

    const currentAvailable =
      String(
        row.Available || ''
      ).toUpperCase() === 'TRUE';

    // ถ้าแถวใหม่เปิดขาย แต่แถวเดิมปิด
    // ให้เก็บแถวใหม่แทน
    if (
      currentAvailable &&
      !seen[key].available
    ) {

      rowsToDelete.push(
        seen[key].sheetRow
      );

      seen[key] = {
        sheetRow:
          index + 2,

        available: true
      };

    } else {

      rowsToDelete.push(
        index + 2
      );

    }

  });

  // ลบจากล่างขึ้นบน
  rowsToDelete
    .sort(function(a, b) {
      return b - a;
    })
    .forEach(function(rowNumber) {

      sheet.deleteRow(
        rowNumber
      );

    });

  SpreadsheetApp.flush();

  return {
    success: true,
    removed:
      rowsToDelete.length
  };

}

/**
 * ==========================================
 * RESET PICKUP SLOTS
 * 08:00 - 17:00 ทุก 20 นาที
 *
 * วิธีใช้:
 * resetPickupSlots20Min(10)
 *
 * 10 = Capacity ต่อ Slot
 * เปลี่ยนเลขได้ตามต้องการ
 * ==========================================
 */
function resetPickupSlots20Min(capacityPerSlot) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetName = CONFIG.SHEETS.SLOTS || 'Pickup_Slots';

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const headers = [
    'Slot_ID',
    'Start_Time',
    'End_Time',
    'Capacity',
    'Active'
  ];

  // ล้างข้อมูลเดิมทั้งหมด
  sheet.clearContents();

  // Header
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  const capacity =
    Number(capacityPerSlot || 10);

  const rows = [];

  let slotNumber = 1;

  // 08:00 = 480 นาที
  const startMinutes = 8 * 60;

  // 17:00 = 1020 นาที
  const endMinutes = 17 * 60;

  for (
    let current = startMinutes;
    current < endMinutes;
    current += 20
  ) {

    const next = current + 20;

    const startHour =
      Math.floor(current / 60);

    const startMinute =
      current % 60;

    const endHour =
      Math.floor(next / 60);

    const endMinute =
      next % 60;

    const startText =
      String(startHour).padStart(2, '0') +
      ':' +
      String(startMinute).padStart(2, '0');

    const endText =
      String(endHour).padStart(2, '0') +
      ':' +
      String(endMinute).padStart(2, '0');

    rows.push([
      'S' + String(slotNumber).padStart(3, '0'),
      startText,
      endText,
      capacity,
      true
    ]);

    slotNumber++;
  }

  if (rows.length > 0) {

    sheet
      .getRange(
        2,
        1,
        rows.length,
        headers.length
      )
      .setValues(rows);

  }

  SpreadsheetApp.flush();

  return {
    success: true,
    count: rows.length,
    capacity: capacity,
    firstSlot: rows[0][1] + ' - ' + rows[0][2],
    lastSlot:
      rows[rows.length - 1][1] +
      ' - ' +
      rows[rows.length - 1][2]
  };
}

function testCustomerDailyDishes() {

  const data = getInitialData();

  Logger.log(
    'APP VERSION = ' +
    data.appVersion
  );

  Logger.log(
    'DAILY DISH COUNT = ' +
    (data.dailyDishes || []).length
  );

  Logger.log(
    JSON.stringify(
      data.dailyDishes,
      null,
      2
    )
  );

}

function testGitHubImageUploadProd() {

  const pngBase64 =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

  const bytes =
    Utilities.base64Decode(
      pngBase64
    );

  const fileName =
    'images/menu/test-prod-' +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyyMMdd-HHmmss'
    ) +
    '-' +
    Utilities.getUuid()
      .replace(/-/g, '')
      .substring(0, 8) +
    '.png';

  const result =
    uploadFileToGitHubProd_(
      fileName,
      bytes,
      'image/png'
    );

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );
}

function sendLineReadyNotification_(
  userId,
  orderId,
  orderDate
) {

  userId =
    String(userId || '').trim();

  orderId =
    String(orderId || '').trim();

  orderDate =
    String(orderDate || '').trim();

  if (!userId) {
    return {
      success: false,
      sent: false,
      reason: 'NO_LINE_USER_ID'
    };
  }

  const token =
    String(
      PropertiesService
        .getScriptProperties()
        .getProperty(
          'LINE_CHANNEL_ACCESS_TOKEN'
        ) || ''
    ).trim();

  if (!token) {
    return {
      success: false,
      sent: false,
      reason:
        'NO_CHANNEL_ACCESS_TOKEN'
    };
  }

  const message =
    '✅ อาหารพร้อมรับแล้วครับ!\n\n' +
    'เลขออเดอร์: ' +
    orderId +
    '\n' +
    'วันที่: ' +
    orderDate +
    '\n\n' +
    'สามารถมารับอาหารที่ร้านได้เลยครับ 😊';

  const response =
    UrlFetchApp.fetch(
      'https://api.line.me/v2/bot/message/push',
      {
        method: 'post',

        contentType:
          'application/json',

        headers: {
          Authorization:
            'Bearer ' + token
        },

        payload:
          JSON.stringify({
            to: userId,
            messages: [
              {
                type: 'text',
                text: message
              }
            ]
          }),

        muteHttpExceptions:
          true
      }
    );

  const status =
    response.getResponseCode();

  const text =
    response.getContentText();

  console.log(
    'LINE READY PUSH STATUS:',
    status
  );

  console.log(
    'LINE READY PUSH RESPONSE:',
    text
  );

  if (
    status < 200 ||
    status >= 300
  ) {
    return {
      success: false,
      sent: false,
      httpStatus: status,
      error: text
    };
  }

  return {
    success: true,
    sent: true,
    httpStatus: status
  };
}

function testReadyNotificationProd() {

  // ==========================================
  // ใส่เลข Order ของ "ออเดอร์ของคุณเอง" ตรงนี้
  // เช่น A012
  // ==========================================
  const TEST_ORDER_ID = 'A005';

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.SHEETS.ORDERS
    );

  if (!sheet) {
    throw new Error(
      'ไม่พบ Sheet Orders'
    );
  }

  const lastRow =
    sheet.getLastRow();

  const lastCol =
    sheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastCol < 4
  ) {
    throw new Error(
      'Orders ยังไม่มีข้อมูล'
    );
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastCol
      )
      .getValues()[0]
      .map(function(header) {
        return String(
          header || ''
        ).trim();
      });

  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastCol
      )
      .getValues();

  const orderIdCol =
    headers.indexOf(
      'Order_ID'
    );

  const dateCol =
    headers.indexOf(
      'Date'
    );

  if (
    orderIdCol < 0
  ) {
    throw new Error(
      'ไม่พบคอลัมน์ Order_ID'
    );
  }

  if (
    dateCol < 0
  ) {
    throw new Error(
      'ไม่พบคอลัมน์ Date'
    );
  }

  // D = LINE User ID
  const userIdCol = 3;

  let target = null;

  rows.forEach(function(row) {

    if (target) return;

    const id =
      String(
        row[orderIdCol] || ''
      ).trim();

    if (
      id === TEST_ORDER_ID
    ) {
      target = row;
    }

  });

  if (!target) {
    throw new Error(
      'ไม่พบ Order ' +
      TEST_ORDER_ID
    );
  }

  const userId =
    String(
      target[userIdCol] || ''
    ).trim();

  if (!userId) {
    throw new Error(
      'Order ' +
      TEST_ORDER_ID +
      ' ไม่มี LINE User ID ในคอลัมน์ D'
    );
  }

  let orderDate = '';

  const rawDate =
    target[dateCol];

  if (
    Object.prototype.toString
      .call(rawDate) ===
    '[object Date]' &&
    !isNaN(
      rawDate.getTime()
    )
  ) {

    orderDate =
      Utilities.formatDate(
        rawDate,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      );

  } else {

    orderDate =
      String(
        rawDate || ''
      ).trim();

  }

  console.log(
    'TEST ORDER =',
    TEST_ORDER_ID
  );

  console.log(
    'ORDER DATE =',
    orderDate
  );

  console.log(
    'LINE USER ID EXISTS =',
    !!userId
  );

  const result =
    sendLineReadyNotification_(
      userId,
      TEST_ORDER_ID,
      orderDate
    );

  console.log(
    'READY NOTIFICATION RESULT =',
    JSON.stringify(
      result,
      null,
      2
    )
  );
}

/**
 * =========================================================
 * PROD - แจ้งเตือนออเดอร์ใหม่เข้า LINE Group ร้าน
 * =========================================================
 */
function sendNewOrderGroupNotificationProd_(
  orderId,
  orderData,
  total,
  detailRows,
  orderDate,
  orderTime
) {

  const props =
    PropertiesService.getScriptProperties();

  const groupId =
    String(
      props.getProperty(
        'SHOP_NOTIFY_GROUP_ID'
      ) || ''
    ).trim();

  if (!groupId) {
    return {
      success: false,
      sent: false,
      reason:
        'NO_SHOP_NOTIFY_GROUP_ID'
    };
  }

  const token =
    String(
      props.getProperty(
        'LINE_CHANNEL_ACCESS_TOKEN'
      ) || ''
    ).trim();

  if (!token) {
    return {
      success: false,
      sent: false,
      reason:
        'NO_CHANNEL_ACCESS_TOKEN'
    };
  }

  const customerName =
    String(
      orderData &&
      orderData.customerName
        ? orderData.customerName
        : ''
    ).trim();

  const customerRemark =
    String(
      orderData &&
      orderData.customerRemark
        ? orderData.customerRemark
        : ''
    ).trim();

  const payment =
    String(
      orderData &&
      orderData.payment
        ? orderData.payment
        : 'CASH'
    ).trim().toUpperCase();

  const pickupType =
    String(
      orderData &&
      orderData.pickupType
        ? orderData.pickupType
        : 'NOW'
    ).trim().toUpperCase();

  const pickupSlot =
    String(
      orderData &&
      orderData.pickupSlot
        ? orderData.pickupSlot
        : ''
    ).trim();

  const paymentTextMap = {
    CASH:
      '💵 เงินสด',

    QR:
      '📱 PromptPay',

    GOVERNMENT:
      '🏛️ โครงการรัฐ / ถุงเงิน'
  };

  const paymentText =
    paymentTextMap[payment] ||
    payment;

  let pickupText =
    'รับเร็วที่สุด';

  if (
    pickupType === 'SCHEDULED' &&
    pickupSlot
  ) {

    pickupText =
      getPickupTimeText_(
        getShopSlotMapProd_(),
        pickupSlot
      ) ||
      pickupSlot;

  }

  const itemLines = [];

  (
    Array.isArray(detailRows)
      ? detailRows
      : []
  ).forEach(
    function(row) {

      /*
       * detailRows:
       * [date, orderId, menuId, menuName,
       *  qty, basePrice, options,
       *  optionsPrice, itemTotal, remark]
       */

      const menuName =
        String(
          row[3] || 'เมนู'
        ).trim();

      const qty =
        Number(
          row[4] || 0
        );

      const options =
        String(
          row[6] || ''
        ).trim();

      const remark =
        String(
          row[9] || ''
        ).trim();

      let line =
        '• ' +
        menuName +
        ' × ' +
        qty;

      if (options) {
        line +=
          ' (' +
          options +
          ')';
      }

      if (remark) {
        line +=
          '\n  📝 ' +
          remark;
      }

      itemLines.push(line);

    }
  );

  if (
    itemLines.length === 0
  ) {
    itemLines.push(
      '• ดูรายละเอียดในหน้า Admin'
    );
  }

  let message =
    '🔔 ออเดอร์ใหม่!\n\n' +

    'เลขออเดอร์: ' +
    orderId +
    '\n' +

    '👤 ลูกค้า: ' +
    (
      customerName ||
      '-'
    ) +
    '\n' +

    (
      orderTime
        ? '🕐 เวลา: ' +
          orderTime +
          '\n'
        : ''
    ) +

    '\n🍛 รายการ\n' +

    itemLines.join('\n') +

    '\n\n💰 ยอดรวม: ' +
    Number(total || 0)
      .toLocaleString('th-TH') +
    ' บาท\n' +

    '💳 ชำระ: ' +
    paymentText +
    '\n' +

    '🕐 รับอาหาร: ' +
    pickupText;

  if (customerRemark) {

    message +=
      '\n\n📝 หมายเหตุลูกค้า: ' +
      customerRemark;

  }

  message +=
    '\n\n👉 กรุณาเปิดหน้า Admin ' +
    'เพื่อรับออเดอร์';

  const response =
    UrlFetchApp.fetch(
      'https://api.line.me/v2/bot/message/push',
      {
        method: 'post',

        contentType:
          'application/json',

        headers: {
          Authorization:
            'Bearer ' +
            token
        },

        payload:
          JSON.stringify({
            to:
              groupId,

            messages: [
              {
                type:
                  'text',

                text:
                  message
              }
            ]
          }),

        muteHttpExceptions:
          true
      }
    );

  const status =
    response.getResponseCode();

  const responseText =
    response.getContentText();

  console.log(
    'PROD SHOP NEW ORDER PUSH STATUS:',
    status
  );

  console.log(
    'PROD SHOP NEW ORDER PUSH RESPONSE:',
    responseText
  );

  if (
    status < 200 ||
    status >= 300
  ) {

    return {
      success: false,
      sent: false,
      httpStatus:
        status,
      error:
        responseText
    };

  }

  return {
    success: true,
    sent: true,
    httpStatus:
      status
  };
}


/**
 * =========================================================
 * PROD - สร้าง Slot Map สำหรับข้อความแจ้งเตือน
 * =========================================================
 */
function getShopSlotMapProd_() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.SHEETS.SLOTS
    );

  if (!sheet) {
    return {};
  }

  const rows =
    getSheetObjects_(
      sheet
    );

  const map = {};

  rows.forEach(
    function(row) {

      const id =
        String(
          row.Slot_ID || ''
        ).trim();

      if (!id) {
        return;
      }

      map[id] = {
        start:
          formatTime_(
            row.Start_Time
          ),

        end:
          formatTime_(
            row.End_Time
          )
      };

    }
  );

  return map;
}


/**
 * =========================================================
 * PROD - ทดสอบส่งข้อความเข้า Group
 * =========================================================
 */
function testShopGroupPushProd() {

  const props =
    PropertiesService.getScriptProperties();

  const groupId =
    String(
      props.getProperty(
        'SHOP_NOTIFY_GROUP_ID'
      ) || ''
    ).trim();

  if (!groupId) {
    throw new Error(
      'ไม่พบ SHOP_NOTIFY_GROUP_ID ใน PROD'
    );
  }

  const token =
    String(
      props.getProperty(
        'LINE_CHANNEL_ACCESS_TOKEN'
      ) || ''
    ).trim();

  if (!token) {
    throw new Error(
      'ไม่พบ LINE_CHANNEL_ACCESS_TOKEN ใน PROD'
    );
  }

  const message =
    '🔔 ทดสอบระบบ Production\n\n' +
    'LINE Group ร้านเชื่อมต่อสำเร็จ ✅\n\n' +
    'พร้อมสำหรับแจ้งเตือนออเดอร์ใหม่';

  const response =
    UrlFetchApp.fetch(
      'https://api.line.me/v2/bot/message/push',
      {
        method: 'post',

        contentType:
          'application/json',

        headers: {
          Authorization:
            'Bearer ' +
            token
        },

        payload:
          JSON.stringify({
            to:
              groupId,

            messages: [
              {
                type:
                  'text',

                text:
                  message
              }
            ]
          }),

        muteHttpExceptions:
          true
      }
    );

  const status =
    response.getResponseCode();

  const text =
    response.getContentText();

  console.log(
    'PROD SHOP GROUP PUSH STATUS:',
    status
  );

  console.log(
    'PROD SHOP GROUP PUSH RESPONSE:',
    text
  );

  if (
    status < 200 ||
    status >= 300
  ) {

    throw new Error(
      'ส่งข้อความเข้า Group PROD ไม่สำเร็จ (' +
      status +
      '): ' +
      text
    );

  }

  console.log(
    'PROD SHOP GROUP PUSH SUCCESS'
  );
}