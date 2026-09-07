/**
 * ============================================================
 * FOOD ORDER SYSTEM CORE BACKEND V1.0 - GOLDEN MASTER
 * ============================================================
 * Derived from production baseline commit 7ceb6c1.
 * All production capabilities are retained.
 * Store-specific values are supplied by Script Properties.
 * Secrets are never embedded in source code.
 * ============================================================
 */

/**
 * Store-aware configuration adapter.
 * Primary source: STORE_CONFIG_JSON
 * Backward-compatible fallback: individual Script Properties.
 */
function getStoreConfig_() {

  const props =
    PropertiesService
      .getScriptProperties();

  const raw =
    String(
      props.getProperty(
        'STORE_CONFIG_JSON'
      ) || ''
    ).trim();

  let config = {};

  if (raw) {

    try {
      config =
        JSON.parse(raw) || {};
    } catch (error) {
      throw new Error(
        'STORE_CONFIG_JSON ไม่ถูกต้อง: ' +
        error.message
      );
    }
  }

  const fallbackKeys = [
    'STORE_ID',
    'SHOP_NAME',
    'PROVIDER',
    'LIFF_ID',
    'LIFF_URL',
    'API_URL',
    'GITHUB_OWNER',
    'GITHUB_REPO',
    'GITHUB_BRANCH',
    'GITHUB_IMAGE_DIR',
    'LINE_LOGIN_CHANNEL_ID',
    'LINE_MESSAGING_CHANNEL_ID'
  ];

  fallbackKeys.forEach(function(key) {

    if (
      config[key] === undefined ||
      config[key] === null ||
      config[key] === ''
    ) {

      const value =
        props.getProperty(key);

      if (
        value !== null &&
        value !== ''
      ) {
        config[key] = value;
      }
    }
  });

  if (!config.PAYMENT) {
    config.PAYMENT = {};
  }

  if (!config.PICKUP) {
    config.PICKUP = {};
  }

  return config;
}


const CONFIG = {

  get LIFF_ID() {
    return String(
      getStoreConfig_().LIFF_ID || ''
    );
  },

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
 * เปิดหน้าเว็บ / API
 */
function doGet(e) {

  if (
    e &&
    e.parameter &&
    e.parameter.action ===
    'getInitialData'
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
    e &&
    e.parameter &&
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

  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle(
      String(
        getStoreConfig_().SHOP_NAME ||
        'FOOD ORDER SYSTEM'
      )
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}


function getInitialData() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const menus =
    getSheetObjects_(
      ss.getSheetByName(
        CONFIG.SHEETS.MENU
      )
    );

  const options =
    getSheetObjects_(
      ss.getSheetByName(
        CONFIG.SHEETS.OPTIONS
      )
    );

  const slots =
    getSheetObjects_(
      ss.getSheetByName(
        CONFIG.SHEETS.SLOTS
      )
    );

  const orders =
    getSheetObjects_(
      ss.getSheetByName(
        CONFIG.SHEETS.ORDERS
      )
    );

  const today =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

  const slotCounts = {};
  const waitingCounts = {};

  orders.forEach(function(order) {

    let orderDate = '';
    const rawDate =
      order.Date;

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

    if (
      orderDate === today
    ) {

      const slot =
        String(
          order.Pickup_Slot || ''
        ).trim();

      const status =
        String(
          order.Status || ''
        ).trim().toUpperCase();

      if (slot) {

        if (
          status !== 'COMPLETED' &&
          status !== 'CANCELLED'
        ) {

          slotCounts[slot] =
            (slotCounts[slot] || 0) + 1;
        }

        if (
          status === 'NEW' ||
          status === 'PREPARING'
        ) {

          waitingCounts[slot] =
            (waitingCounts[slot] || 0) + 1;
        }
      }
    }
  });

  const dailySheet =
    ss.getSheetByName(
      'Daily_Dishes'
    );

  const dailyRows =
    dailySheet
      ? getSheetObjects_(
          dailySheet
        )
      : [];

  const dailyDishes =
    dailyRows
      .filter(function(dish) {
        return (
          normalizeDateValue_(
            dish.Dish_Date
          ) === today &&
          String(
            dish.Available || ''
          ).toUpperCase() ===
          'TRUE'
        );
      })
      .map(function(dish) {
        return {
          id: String(
            dish.Dish_ID || ''
          ),
          name: String(
            dish.Dish_Name || ''
          ),
          description: String(
            dish.Description || ''
          ),
          image: String(
            dish.Image_URL || ''
          )
        };
      });

  const specialSheet =
    ss.getSheetByName(
      'Special_Menus'
    );

  const specialRows =
    specialSheet
      ? getSheetObjects_(
          specialSheet
        )
      : [];

  const specialMenus =
    specialRows
      .filter(function(menu) {

        const start =
          normalizeDateValue_(
            menu.Start_Date ||
            today
          );

        const end =
          normalizeDateValue_(
            menu.End_Date ||
            today
          );

        return (
          String(
            menu.Available || ''
          ).toUpperCase() ===
          'TRUE' &&
          today >= start &&
          today <= end
        );
      })
      .map(function(menu) {
        return {
          id: String(
            menu.Special_ID || ''
          ),
          name: String(
            menu.Menu_Name || ''
          ),
          description: String(
            menu.Description || ''
          ),
          price: Number(
            menu.Price || 0
          ),
          category: '⭐ เมนูพิเศษ',
          image: String(
            menu.Image_URL || ''
          ),
          menuType: 'SPECIAL'
        };
      });

  const mainMenus =
    menus
      .filter(function(menu) {
        return String(
          menu.Available || ''
        ).toUpperCase() === 'TRUE';
      })
      .map(function(menu) {
        return {
          id: String(
            menu.Menu_ID || ''
          ),
          name: String(
            menu.Menu_Name || ''
          ),
          description: String(
            menu.Description || ''
          ),
          price: Number(
            menu.Price || 0
          ),
          category: String(
            menu.Category || ''
          ),
          image: String(
            menu.Image_URL || ''
          ),
          menuType: 'MAIN'
        };
      });

  const visibleOptions =
    options
      .filter(function(option) {
        return String(
          option.Available || ''
        ).toUpperCase() === 'TRUE';
      })
      .map(function(option) {
        return {
          menuId: String(
            option.Menu_ID || ''
          ),
          name: String(
            option.Option_Name || ''
          ),
          price: Number(
            option.Option_Price || 0
          ),
          group: String(
            option.Option_Group || ''
          ),
          minSelect: Number(
            option.Min_Select || 0
          ),
          maxSelect: Number(
            option.Max_Select || 0
          )
        };
      });

  const curryConfig = {
    M022: 1,
    M023: 2,
    M024: 3
  };

  Object.keys(
    curryConfig
  ).forEach(function(menuId) {

    const required =
      curryConfig[menuId];

    dailyDishes.forEach(
      function(dish) {

        visibleOptions.push({
          menuId: menuId,
          name: dish.name,
          price: 0,
          group: 'เลือกกับข้าว',
          minSelect: required,
          maxSelect: required,
          dynamicDailyDish: true
        });
      }
    );
  });

  return {
    appVersion:
      'BACKEND-V4-MENU-SYSTEM',
    liffId:
      CONFIG.LIFF_ID,
    payment:
      getPaymentSettings_(),
    menus:
      mainMenus.concat(
        specialMenus
      ),
    options:
      visibleOptions,
    dailyDishes:
      dailyDishes,
    specialMenus:
      specialMenus,
    slots:
      slots
        .filter(function(slot) {
          return String(
            slot.Active || ''
          ).toUpperCase() ===
          'TRUE';
        })
        .map(function(slot) {

          const id = String(
            slot.Slot_ID || ''
          );

          const capacity =
            Number(
              slot.Capacity || 0
            );

          const used =
            slotCounts[id] || 0;

          return {
            id: id,
            start:
              formatTime_(
                slot.Start_Time
              ),
            end:
              formatTime_(
                slot.End_Time
              ),
            capacity: capacity,
            used: used,
            waiting:
              waitingCounts[id] || 0,
            waitingCount:
              waitingCounts[id] || 0,
            remaining:
              Math.max(
                capacity - used,
                0
              )
          };
        })
  };
}


function createOrder(orderData) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(15000);

  try {

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

    const menus =
      getSheetObjects_(
        menuSheet
      );

    const options =
      getSheetObjects_(
        optionSheet
      );

    const dailySheet =
      ss.getSheetByName(
        'Daily_Dishes'
      );

    const menuSystemToday =
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      );

    const dailyDishes =
      dailySheet
        ? getSheetObjects_(
            dailySheet
          ).filter(function(dish) {
            return (
              normalizeDateValue_(
                dish.Dish_Date
              ) === menuSystemToday &&
              String(
                dish.Available || ''
              ).toUpperCase() ===
              'TRUE'
            );
          })
        : [];

    const specialSheet =
      ss.getSheetByName(
        'Special_Menus'
      );

    const specialMenus =
      specialSheet
        ? getSheetObjects_(
            specialSheet
          ).filter(function(menu) {

            const startDate =
              String(
                menu.Start_Date ||
                menuSystemToday
              ).trim();

            const endDate =
              String(
                menu.End_Date ||
                menuSystemToday
              ).trim();

            return (
              String(
                menu.Available || ''
              ).toUpperCase() ===
              'TRUE' &&
              menuSystemToday >=
                startDate &&
              menuSystemToday <=
                endDate
            );
          })
        : [];

    const menuMap = {};

    menus.forEach(function(menu) {

      if (
        String(
          menu.Available || ''
        ).toUpperCase() === 'TRUE'
      ) {

        menuMap[
          String(menu.Menu_ID)
        ] = {
          id: String(
            menu.Menu_ID
          ),
          name: String(
            menu.Menu_Name
          ),
          price: Number(
            menu.Price || 0
          )
        };
      }
    });

    specialMenus.forEach(
      function(menu) {

        menuMap[
          String(menu.Special_ID)
        ] = {
          id: String(
            menu.Special_ID
          ),
          name: String(
            menu.Menu_Name
          ),
          price: Number(
            menu.Price || 0
          )
        };
      }
    );

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
          curryRequiredMap[
            String(item.menuId)
          ] || 0;

        if (curryRequired > 0) {

          const selectedNames =
            (item.options || [])
              .map(function(selected) {
                return String(
                  selected.name || ''
                ).trim();
              })
              .filter(Boolean);

          const uniqueNames =
            selectedNames.filter(
              function(name, index, arr) {
                return (
                  arr.indexOf(name) ===
                  index
                );
              }
            );

          if (
            uniqueNames.length !==
            curryRequired
          ) {
            throw new Error(
              'กรุณาเลือกกับข้าว ' +
              curryRequired +
              ' อย่าง'
            );
          }
        }

        (item.options || [])
          .forEach(function(selected) {

            const found =
              options.find(
                function(option) {
                  return (
                    String(
                      option.Menu_ID
                    ) ===
                      String(
                        item.menuId
                      ) &&
                    String(
                      option.Option_Name
                    ) ===
                      String(
                        selected.name
                      ) &&
                    String(
                      option.Available
                    ).toUpperCase() ===
                      'TRUE'
                  );
                }
              );

            const isDailyDishOption =
              [
                'M022',
                'M023',
                'M024'
              ].indexOf(
                String(item.menuId)
              ) >= 0 &&
              dailyDishes.some(
                function(dish) {
                  return (
                    String(
                      dish.Dish_Name
                    ) ===
                    String(
                      selected.name
                    )
                  );
                }
              );

            if (
              found ||
              isDailyDishOption
            ) {

              optionTotal +=
                found
                  ? Number(
                      found.Option_Price ||
                      0
                    )
                  : 0;

              optionNames.push(
                String(
                  selected.name
                )
              );
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
          String(
            item.remark || ''
          )
        ]);
      }
    );

    const pickupType =
      String(
        orderData.pickupType ||
        'NOW'
      );

    let pickupSlot = '';

    if (
      pickupType === 'SCHEDULED'
    ) {

      pickupSlot =
        String(
          orderData.pickupSlot ||
          ''
        );

      if (!pickupSlot) {
        throw new Error(
          'กรุณาเลือกช่วงเวลารับ'
        );
      }

      const slots =
        getSheetObjects_(
          slotSheet
        );

      const selectedSlot =
        slots.find(function(slot) {
          return (
            String(
              slot.Slot_ID
            ) === pickupSlot &&
            String(
              slot.Active
            ).toUpperCase() ===
            'TRUE'
          );
        });

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
        getSheetObjects_(
          orderSheet
        );

      let used = 0;

      orders.forEach(function(order) {

        let orderDate = '';
        const rawDate =
          order.Date;

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

        const status =
          String(
            order.Status || ''
          ).trim().toUpperCase();

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
      });

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

      const values =
        orderSheet
          .getRange(
            2,
            1,
            lastRow - 1,
            2
          )
          .getValues();

      values.forEach(function(row) {

        const rawOrderId =
          row[0];

        const rawDate =
          row[1];

        let rowDate = '';

        if (
          Object.prototype.toString
            .call(rawDate) ===
            '[object Date]' &&
          !isNaN(
            rawDate.getTime()
          )
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

        if (
          rowDate !== today
        ) {
          return;
        }

        const id =
          String(
            rawOrderId || ''
          ).trim().toUpperCase();

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
      });
    }

    const orderId =
      'A' +
      String(
        maxNumber + 1
      ).padStart(3, '0');

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

    const now =
      new Date();

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
        orderData.payment ||
        'CASH'
      );

    const customerName =
      String(
        orderData.customerName ||
        ''
      );

    const customerRemark =
      String(
        orderData.customerRemark ||
        ''
      );

    orderSheet.appendRow([
      orderId,
      date,
      time,
      String(
        orderData.userId
      ),
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
        detailSheet.appendRow(
          row
        );
      }
    );

    saveCustomer_(
      customerSheet,
      String(
        orderData.userId
      ),
      customerName
    );

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
        'SHOP NEW ORDER NOTIFICATION ERROR:',
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
      orderId: orderId,
      total: total,
      pickupType: pickupType,
      pickupSlot: pickupSlot,
      shopNotification:
        shopNotification
    };

  } finally {
    lock.releaseLock();
  }
}


function saveCustomer_(
  sheet,
  userId,
  name
) {

  const values =
    sheet.getDataRange()
      .getValues();

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(
        values[i][0] || ''
      ).trim() === userId
    ) {

      sheet
        .getRange(
          i + 1,
          2
        )
        .setValue(name);

      return;
    }
  }

  sheet.appendRow([
    userId,
    name,
    name,
    ''
  ]);
}


function getSheetObjects_(
  sheet
) {

  if (!sheet) {
    return [];
  }

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {
    return [];
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0]
      .map(function(header) {
        return String(
          header || ''
        ).trim();
      });

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getValues();

  return values.map(
    function(row) {

      const obj = {};

      headers.forEach(
        function(header, index) {
          if (header) {
            obj[header] =
              row[index];
          }
        }
      );

      return obj;
    }
  );
}


function normalizeDateValue_(
  value
) {

  if (
    Object.prototype.toString
      .call(value) ===
    '[object Date]' &&
    !isNaN(
      value.getTime()
    )
  ) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  const text =
    String(
      value || ''
    ).trim();

  if (!text) {
    return '';
  }

  const match =
    text.match(
      /^(\d{4})[-\/]?(\d{2})[-\/]?(\d{2})/
    );

  if (match) {
    return (
      match[1] +
      '-' +
      match[2] +
      '-' +
      match[3]
    );
  }

  return text;
}


function getHeaders_(
  sheet
) {

  if (!sheet) {
    return [];
  }

  const lastColumn =
    sheet.getLastColumn();

  if (lastColumn < 1) {
    return [];
  }

  return sheet
    .getRange(
      1,
      1,
      1,
      lastColumn
    )
    .getValues()[0]
    .map(function(header) {
      return String(
        header || ''
      ).trim();
    });
}


function formatTime_(
  value
) {

  if (
    Object.prototype.toString
      .call(value) ===
    '[object Date]' &&
    !isNaN(
      value.getTime()
    )
  ) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'HH:mm'
    );
  }

  const text =
    String(
      value || ''
    ).trim();

  const match =
    text.match(
      /^(\d{1,2}):(\d{2})/
    );

  if (match) {
    return (
      String(
        Number(match[1])
      ).padStart(2, '0') +
      ':' +
      match[2]
    );
  }

  return text;
}


function testPickupSlots() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.SHEETS.SLOTS
    );

  if (!sheet) {
    throw new Error(
      'ไม่พบ Sheet Pickup_Slots'
    );
  }

  Logger.log(
    JSON.stringify(
      getSheetObjects_(
        sheet
      ),
      null,
      2
    )
  );
}


function doGetDailyDishLibrary_(
  e
) {

  requireAdminAuth_(
    e &&
    e.parameter
      ? e.parameter.adminToken
      : ''
  );

  return jsonResponse_(
    getDailyDishLibrary()
  );
}


function doPost(e) {

  try {

    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {

      throw new Error(
        'ไม่พบ POST body'
      );
    }

    const data =
      JSON.parse(
        e.postData.contents
      );

    const result =
      createOrder(
        data
      );

    return jsonResponse_(
      result
    );

  } catch (error) {

    return jsonResponse_({
      success: false,
      error:
        error.message ||
        String(error)
    });
  }
}


function jsonResponse_(
  data
) {

  return ContentService
    .createTextOutput(
      JSON.stringify(
        data
      )
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
    return {
      success: false,
      orders: [],
      error:
        'ไม่พบ LINE User ID'
    };
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

  const orders =
    getSheetObjects_(
      orderSheet
    ).filter(function(order) {
      return String(
        order.LINE_User_ID || ''
      ).trim() === id;
    });

  const details =
    getSheetObjects_(
      detailSheet
    );

  const mapped =
    orders.map(function(order) {

      const orderId = String(
        order.Order_ID || ''
      );

      const items =
        details
          .filter(function(detail) {
            return String(
              detail.Order_ID || ''
            ) === orderId;
          })
          .map(function(detail) {
            return {
              menuId: String(
                detail.Menu_ID || ''
              ),
              menuName: String(
                detail.Menu_Name || ''
              ),
              qty: Number(
                detail.Qty || 0
              ),
              unitPrice: Number(
                detail.Unit_Price || 0
              ),
              options: String(
                detail.Options || ''
              ),
              optionTotal: Number(
                detail.Option_Total || 0
              ),
              itemTotal: Number(
                detail.Item_Total || 0
              ),
              remark: String(
                detail.Remark || ''
              )
            };
          });

      return {
        orderId: orderId,
        date: String(
          order.Date || ''
        ),
        time: String(
          order.Time || ''
        ),
        customerName: String(
          order.Customer_Name || ''
        ),
        pickupType: String(
          order.Pickup_Type || ''
        ),
        pickupSlot: String(
          order.Pickup_Slot || ''
        ),
        total: Number(
          order.Total || 0
        ),
        payment: String(
          order.Payment || ''
        ),
        paymentStatus: String(
          order.Payment_Status || ''
        ),
        status: String(
          order.Status || ''
        ),
        customerRemark: String(
          order.Customer_Remark || ''
        ),
        createdAt: String(
          order.Created_At || ''
        ),
        items: items
      };
    });

  mapped.sort(function(a, b) {
    return String(
      b.createdAt
    ).localeCompare(
      String(a.createdAt)
    );
  });

  return {
    success: true,
    orders: mapped
  };
}


function getOrders() {

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

  const orders =
    getSheetObjects_(
      orderSheet
    );

  const details =
    getSheetObjects_(
      detailSheet
    );

  return {
    success: true,
    orders: orders.map(
      function(order) {

        const orderId = String(
          order.Order_ID || ''
        );

        return Object.assign(
          {},
          order,
          {
            items:
              details.filter(
                function(detail) {
                  return String(
                    detail.Order_ID || ''
                  ) === orderId;
                }
              )
          }
        );
      }
    )
  };
}


function formatTimeValue_(
  value
) {
  return formatTime_(
    value
  );
}


function getPickupTimeText_(
  slotMap,
  slotId
) {

  const id =
    String(
      slotId || ''
    ).trim();

  if (!id || !slotMap) {
    return '';
  }

  const slot =
    slotMap[id];

  if (!slot) {
    return '';
  }

  return (
    String(slot.start || '') +
    ' - ' +
    String(slot.end || '')
  );
}


function updateOrderStatus(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  const orderId =
    String(
      data.orderId || ''
    ).trim();

  const nextStatus =
    String(
      data.status || ''
    ).trim().toUpperCase();

  if (!orderId) {
    throw new Error(
      'ไม่พบ Order ID'
    );
  }

  if (
    [
      'NEW',
      'PREPARING',
      'READY',
      'COMPLETED',
      'CANCELLED'
    ].indexOf(
      nextStatus
    ) < 0
  ) {
    throw new Error(
      'สถานะ Order ไม่ถูกต้อง'
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.SHEETS.ORDERS
    );

  const headers =
    getHeaders_(sheet);

  const idCol =
    headers.indexOf(
      'Order_ID'
    );

  const statusCol =
    headers.indexOf(
      'Status'
    );

  if (
    idCol < 0 ||
    statusCol < 0
  ) {
    throw new Error(
      'Orders ขาดคอลัมน์ Order_ID หรือ Status'
    );
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error(
      'Orders ยังไม่มีข้อมูล'
    );
  }

  const values =
    sheet
      .getRange(
        2,
        idCol + 1,
        lastRow - 1,
        1
      )
      .getValues();

  let rowNumber = -1;

  values.forEach(
    function(row, index) {
      if (
        String(
          row[0] || ''
        ).trim() === orderId
      ) {
        rowNumber =
          index + 2;
      }
    }
  );

  if (
    rowNumber < 0
  ) {
    throw new Error(
      'ไม่พบ Order: ' +
      orderId
    );
  }

  sheet
    .getRange(
      rowNumber,
      statusCol + 1
    )
    .setValue(
      nextStatus
    );

  if (
    nextStatus === 'READY'
  ) {

    const dateCol =
      headers.indexOf(
        'Date'
      );

    const userIdCol =
      headers.indexOf(
        'LINE_User_ID'
      );

    const rowValues =
      sheet
        .getRange(
          rowNumber,
          1,
          1,
          headers.length
        )
        .getValues()[0];

    const userId =
      userIdCol >= 0
        ? String(
            rowValues[userIdCol] ||
            ''
          ).trim()
        : '';

    let orderDate = '';

    if (dateCol >= 0) {
      orderDate =
        normalizeDateValue_(
          rowValues[dateCol]
        );
    }

    try {
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
    }
  }

  return {
    success: true,
    orderId: orderId,
    status: nextStatus
  };
}


function getPaymentSettings_() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'PAYMENT_SETTINGS'
    );

  if (!sheet) {
    return {
      cashActive: true,
      promptPayActive: false,
      governmentActive: false,
      promptPayName: '',
      promptPayNumber: '',
      promptPayQrUrl: '',
      governmentName: '',
      governmentNote: ''
    };
  }

  const rows =
    getSheetObjects_(
      sheet
    );

  const map = {};

  rows.forEach(function(row) {
    map[
      String(
        row.Key || ''
      ).trim()
    ] = String(
      row.Value || ''
    ).trim();
  });

  return {
    cashActive:
      String(
        map.Cash_Active ||
        'TRUE'
      ).toUpperCase() === 'TRUE',
    promptPayActive:
      String(
        map.PromptPay_Active ||
        'FALSE'
      ).toUpperCase() === 'TRUE',
    governmentActive:
      String(
        map.Government_Active ||
        'FALSE'
      ).toUpperCase() === 'TRUE',
    promptPayName:
      String(
        map.PromptPay_Name || ''
      ),
    promptPayNumber:
      String(
        map.PromptPay_Number || ''
      ),
    promptPayQrUrl:
      String(
        map.PromptPay_QR_URL || ''
      ),
    governmentName:
      String(
        map.Government_Name || ''
      ),
    governmentNote:
      String(
        map.Government_Note || ''
      )
  };
}


function getPaymentQRData_() {

  const settings =
    getPaymentSettings_();

  return {
    success: true,
    enabled:
      settings.promptPayActive,
    name:
      settings.promptPayName,
    number:
      settings.promptPayNumber,
    qrUrl:
      settings.promptPayQrUrl
  };
}


function testPaymentStatus() {
  Logger.log(
    JSON.stringify(
      getPaymentSettings_(),
      null,
      2
    )
  );
}


function testPaymentQR() {
  Logger.log(
    JSON.stringify(
      getPaymentQRData_(),
      null,
      2
    )
  );
}


function getAdminSettings() {

  const config =
    getStoreConfig_();

  return {
    success: true,
    storeId: String(
      config.STORE_ID || ''
    ),
    shopName: String(
      config.SHOP_NAME ||
      'FOOD ORDER SYSTEM'
    ),
    provider: String(
      config.PROVIDER || ''
    )
  };
}


function findMenuHeaderColumn_(
  headers,
  candidates
) {

  for (
    let i = 0;
    i < candidates.length;
    i++
  ) {

    const index =
      headers.indexOf(
        candidates[i]
      );

    if (index >= 0) {
      return index;
    }
  }

  return -1;
}


function saveMenu(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  const menu =
    data.menu || {};

  const id =
    String(
      menu.id ||
      menu.menuId ||
      ''
    ).trim();

  if (!id) {
    throw new Error(
      'ไม่พบ Menu ID'
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.SHEETS.MENU
    );

  const headers =
    getHeaders_(sheet);

  const idCol =
    headers.indexOf(
      'Menu_ID'
    );

  if (idCol < 0) {
    throw new Error(
      'Menu ไม่มีคอลัมน์ Menu_ID'
    );
  }

  const row = [
    id,
    String(
      menu.name ||
      menu.Menu_Name ||
      ''
    ),
    String(
      menu.description ||
      menu.Description ||
      ''
    ),
    Number(
      menu.price ||
      menu.Price ||
      0
    ),
    String(
      menu.category ||
      menu.Category ||
      ''
    ),
    String(
      menu.image ||
      menu.Image_URL ||
      ''
    ),
    menu.available === undefined
      ? true
      : Boolean(
          menu.available
        )
  ];

  const values =
    sheet
      .getDataRange()
      .getValues();

  let targetRow = -1;

  values.slice(1).forEach(
    function(existing, index) {
      if (
        String(
          existing[idCol] || ''
        ).trim() === id
      ) {
        targetRow =
          index + 2;
      }
    }
  );

  if (targetRow < 0) {
    sheet.appendRow(row);
  } else {
    sheet
      .getRange(
        targetRow,
        1,
        1,
        row.length
      )
      .setValues([row]);
  }

  return {
    success: true,
    menuId: id
  };
}


function toggleMenuAvailable(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  return setAvailabilityById_(
    CONFIG.SHEETS.MENU,
    'Menu_ID',
    data.menuId,
    'Available',
    data.available
  );
}


function createMenu(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  const menu =
    data.menu || {};

  const name =
    String(
      menu.name ||
      menu.Menu_Name ||
      ''
    ).trim();

  if (!name) {
    throw new Error(
      'กรุณาระบุชื่อเมนู'
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.SHEETS.MENU
    );

  const id =
    String(
      menu.id ||
      menu.menuId ||
      ''
    ).trim() ||
    'M' +
    String(
      Math.floor(
        Math.random() * 9000
      ) + 1000
    );

  sheet.appendRow([
    id,
    name,
    String(
      menu.description || ''
    ),
    Number(
      menu.price || 0
    ),
    String(
      menu.category || ''
    ),
    String(
      menu.image || ''
    ),
    true
  ]);

  return {
    success: true,
    menuId: id
  };
}


function savePaymentSettings(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  const settings =
    data.settings || {};

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'PAYMENT_SETTINGS'
    );

  if (!sheet) {
    throw new Error(
      'ไม่พบ PAYMENT_SETTINGS'
    );
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        Math.max(
          sheet.getLastRow() - 1,
          0
        ),
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

  values.forEach(
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


function uploadAdminImage(
  image
) {

  if (!image) {
    throw new Error(
      'ไม่พบรูปภาพ'
    );
  }

  const dataUrl =
    String(
      image.dataUrl || ''
    ).trim();

  if (!dataUrl) {
    throw new Error(
      'ไม่พบข้อมูลรูปภาพ'
    );
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
    String(
      match[1]
    ).toLowerCase();

  const supportedTypes = [
    'image/jpeg',
    'image/png',
    'image/webp'
  ];

  if (
    supportedTypes.indexOf(
      mimeType
    ) === -1
  ) {
    throw new Error(
      'รองรับเฉพาะ JPG, PNG และ WebP'
    );
  }

  const bytes =
    Utilities.base64Decode(
      match[2]
    );

  if (
    !bytes ||
    !bytes.length
  ) {
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
      .replace(
        /-/g,
        ''
      )
      .substring(
        0,
        8
      ) +
    '.' +
    extension;

  const repoPath =
    String(
      GITHUB_PROD_IMAGE_DIR_ ||
      'images/menu'
    ).replace(
      /\/$/,
      ''
    ) +
    '/' +
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
  getStoreProperty_(
    'GITHUB_OWNER',
    ''
  );

const GITHUB_PROD_REPO_ =
  getStoreProperty_(
    'GITHUB_REPO',
    ''
  );

const GITHUB_PROD_BRANCH_ =
  getStoreProperty_(
    'GITHUB_BRANCH',
    'main'
  );

const GITHUB_PROD_IMAGE_DIR_ =
  getStoreProperty_(
    'GITHUB_IMAGE_DIR',
    'images/menu'
  );


function getStoreProperty_(
  key,
  fallback
) {

  const value = String(
    PropertiesService
      .getScriptProperties()
      .getProperty(
        key
      ) || ''
  ).trim();

  return (
    value ||
    (
      fallback === undefined
        ? ''
        : fallback
    )
  );
}


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

  if (
    !GITHUB_PROD_OWNER_ ||
    !GITHUB_PROD_REPO_
  ) {
    throw new Error(
      'ยังไม่ได้ตั้งค่า GITHUB_OWNER / GITHUB_REPO ของ Store Instance'
    );
  }

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
        return encodeURIComponent(
          part
        );
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
            '2022-11-28'
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


const ADMIN_AUTH_PIN_KEY_ =
  'ADMIN_AUTH_PIN';

const ADMIN_AUTH_SESSIONS_KEY_ =
  'ADMIN_AUTH_SESSIONS';


function generateInitialAdminPin_() {

  let pin = '';

  for (
    let i = 0;
    i < 4;
    i++
  ) {

    pin += String(
      Math.floor(
        Math.random() * 10
      )
    );
  }

  return pin;
}


function setupAdminSecurity() {

  const props =
    PropertiesService
      .getScriptProperties();

  const existing =
    props.getProperty(
      ADMIN_AUTH_PIN_KEY_
    );

  if (!existing) {

    const initialPin =
      generateInitialAdminPin_();

    props.setProperty(
      ADMIN_AUTH_PIN_KEY_,
      initialPin
    );

    return {
      success: true,
      message:
        'ตั้งค่า Admin PIN สำเร็จ',
      adminPin:
        initialPin
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

  Object.keys(
    sessions
  ).forEach(function(key) {

    if (
      !sessions[key] ||
      Number(
        sessions[key]
      ) <= Date.now()
    ) {
      delete sessions[key];
    }
  });

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
    token: token,
    expiresAt: expiresAt
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
      JSON.parse(raw);
  } catch (error) {
    throw new Error(
      'Admin Session ไม่ถูกต้อง'
    );
  }

  const expiresAt =
    Number(
      sessions[authToken] ||
      0
    );

  if (
    !expiresAt ||
    expiresAt <= Date.now()
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
      JSON.parse(raw);
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


function getMenuSystemData() {

  const data =
    getInitialData();

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const menuSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.MENU
    );

  const optionSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.OPTIONS
    );

  const dailySheet =
    ss.getSheetByName(
      'Daily_Dishes'
    );

  const specialSheet =
    ss.getSheetByName(
      'Special_Menus'
    );

  const menus =
    menuSheet
      ? getSheetObjects_(
          menuSheet
        ).map(function(menu) {
          return {
            id: String(
              menu.Menu_ID || ''
            ),
            name: String(
              menu.Menu_Name || ''
            ),
            description: String(
              menu.Description || ''
            ),
            price: Number(
              menu.Price || 0
            ),
            category: String(
              menu.Category || ''
            ),
            image: String(
              menu.Image_URL || ''
            ),
            available:
              String(
                menu.Available || ''
              ).toUpperCase() ===
              'TRUE'
          };
        })
      : [];

  const options =
    optionSheet
      ? getSheetObjects_(
          optionSheet
        ).map(function(option) {
          return {
            menuId: String(
              option.Menu_ID || ''
            ),
            group: String(
              option.Option_Group || ''
            ),
            name: String(
              option.Option_Name || ''
            ),
            price: Number(
              option.Option_Price || 0
            ),
            available:
              String(
                option.Available || ''
              ).toUpperCase() ===
              'TRUE',
            minSelect: Number(
              option.Min_Select || 0
            ),
            maxSelect: Number(
              option.Max_Select || 0
            )
          };
        })
      : [];

  const dailyDishes =
    dailySheet
      ? getSheetObjects_(
          dailySheet
        ).map(function(dish) {
          return {
            id: String(
              dish.Dish_ID || ''
            ),
            date:
              normalizeDateValue_(
                dish.Dish_Date
              ),
            name: String(
              dish.Dish_Name || ''
            ),
            description: String(
              dish.Description || ''
            ),
            image: String(
              dish.Image_URL || ''
            ),
            available:
              String(
                dish.Available || ''
              ).toUpperCase() ===
              'TRUE'
          };
        })
      : [];

  const specialMenus =
    specialSheet
      ? getSheetObjects_(
          specialSheet
        ).map(function(menu) {
          return {
            id: String(
              menu.Special_ID || ''
            ),
            startDate: String(
              menu.Start_Date || ''
            ),
            endDate: String(
              menu.End_Date || ''
            ),
            name: String(
              menu.Menu_Name || ''
            ),
            description: String(
              menu.Description || ''
            ),
            price: Number(
              menu.Price || 0
            ),
            image: String(
              menu.Image_URL || ''
            ),
            available:
              String(
                menu.Available || ''
              ).toUpperCase() ===
              'TRUE'
          };
        })
      : [];

  return {
    success: true,
    today:
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      ),
    data: data,
    menus: menus,
    options: options,
    dailyDishes: dailyDishes,
    specialMenus: specialMenus
  };
}


function ensureSheetWithHeaders_(
  name,
  headers
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(name);

  if (!sheet) {
    sheet =
      ss.insertSheet(name);
  }

  if (
    sheet.getLastRow() === 0
  ) {
    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setValues([
        headers
      ]);
  }

  return sheet;
}


function backupSheetIfExists_(
  sheetName
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      sheetName
    );

  if (!sheet) {
    return '';
  }

  const stamp =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyyMMdd_HHmmss'
    );

  let name =
    sheetName +
    '_BACKUP_' +
    stamp;

  let i = 1;

  while (
    ss.getSheetByName(name)
  ) {
    name =
      sheetName +
      '_BACKUP_' +
      stamp +
      '_' +
      i++;
  }

  sheet
    .copyTo(ss)
    .setName(name);

  return name;
}


function installRecommendedMenuSystem(
  confirm
) {

  if (
    String(
      confirm
    ).toUpperCase() !==
    'YES'
  ) {
    throw new Error(
      'ต้องส่ง confirm = YES เพื่อยืนยันการติดตั้งชุดเมนูใหม่'
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  backupSheetIfExists_(
    CONFIG.SHEETS.MENU
  );

  backupSheetIfExists_(
    CONFIG.SHEETS.OPTIONS
  );

  const menuSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.MENU
    );

  const optionSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.OPTIONS
    );

  const menuHeaders = [
    'Menu_ID',
    'Menu_Name',
    'Description',
    'Price',
    'Category',
    'Image_URL',
    'Available'
  ];

  const optionHeaders = [
    'Menu_ID',
    'Option_Group',
    'Option_Name',
    'Option_Price',
    'Available',
    'Min_Select',
    'Max_Select'
  ];

  menuSheet.clearContents();
  optionSheet.clearContents();

  menuSheet
    .getRange(
      1,
      1,
      1,
      menuHeaders.length
    )
    .setValues([
      menuHeaders
    ]);

  optionSheet
    .getRange(
      1,
      1,
      1,
      optionHeaders.length
    )
    .setValues([
      optionHeaders
    ]);

  return {
    success: true,
    message:
      'ติดตั้ง Menu System V2 แล้ว'
  };
}


function getDailyDishLibrary() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Daily_Dish_Library'
    );

  return {
    success: true,
    dishes:
      getSheetObjects_(
        sheet
      )
  };
}


function addDailyDishesFromLibrary(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  const ids =
    Array.isArray(
      data.libraryIds
    )
      ? data.libraryIds
      : [];

  const date =
    String(
      data.date || ''
    ).trim();

  if (!date) {
    throw new Error(
      'กรุณาระบุวันที่'
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const librarySheet =
    ss.getSheetByName(
      'Daily_Dish_Library'
    );

  const dailySheet =
    ss.getSheetByName(
      'Daily_Dishes'
    );

  const libraryRows =
    getSheetObjects_(
      librarySheet
    );

  const selected =
    libraryRows.filter(
      function(row) {
        return ids.indexOf(
          String(
            row.Library_ID || ''
          )
        ) >= 0;
      }
    );

  selected.forEach(
    function(row) {

      const dishId =
        'D' +
        Utilities.getUuid()
          .replace(
            /-/g,
            ''
          )
          .substring(0, 8)
          .toUpperCase();

      dailySheet.appendRow([
        dishId,
        date,
        String(
          row.Dish_Name || ''
        ),
        String(
          row.Description || ''
        ),
        String(
          row.Image_URL || ''
        ),
        true
      ]);
    }
  );

  return {
    success: true,
    count: selected.length
  };
}


function createDailyDish(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  const dish =
    data.dish || {};

  const id =
    String(
      dish.id ||
      dish.Dish_ID ||
      ''
    ).trim() ||
    'D' +
    Utilities.getUuid()
      .replace(
        /-/g,
        ''
      )
      .substring(0, 8)
      .toUpperCase();

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Daily_Dishes'
    );

  sheet.appendRow([
    id,
    String(
      dish.date ||
      dish.Dish_Date ||
      ''
    ),
    String(
      dish.name ||
      dish.Dish_Name ||
      ''
    ),
    String(
      dish.description ||
      dish.Description ||
      ''
    ),
    String(
      dish.image ||
      dish.Image_URL ||
      ''
    ),
    dish.available === undefined
      ? true
      : Boolean(
          dish.available
        )
  ]);

  return {
    success: true,
    dishId: id
  };
}


function saveDailyDish(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  return updateSimpleRow_(
    'Daily_Dishes',
    'Dish_ID',
    data.dishId ||
    (data.dish || {}).id,
    data.dish || {}
  );
}


function toggleDailyDish(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  return setAvailabilityById_(
    'Daily_Dishes',
    'Dish_ID',
    data.dishId,
    'Available',
    data.available
  );
}


function deleteDailyDish(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  return deleteById_(
    'Daily_Dishes',
    'Dish_ID',
    data.dishId
  );
}


function createSpecialMenu(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  const menu =
    data.menu || {};

  const id =
    String(
      menu.id ||
      menu.Special_ID ||
      ''
    ).trim() ||
    'SM' +
    Utilities.getUuid()
      .replace(
        /-/g,
        ''
      )
      .substring(0, 8)
      .toUpperCase();

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Special_Menus'
    );

  sheet.appendRow([
    id,
    String(
      menu.startDate ||
      menu.Start_Date ||
      ''
    ),
    String(
      menu.endDate ||
      menu.End_Date ||
      ''
    ),
    String(
      menu.name ||
      menu.Menu_Name ||
      ''
    ),
    String(
      menu.description ||
      menu.Description ||
      ''
    ),
    Number(
      menu.price ||
      0
    ),
    String(
      menu.image ||
      menu.Image_URL ||
      ''
    ),
    menu.available === undefined
      ? true
      : Boolean(
          menu.available
        )
  ]);

  return {
    success: true,
    specialId: id
  };
}


function saveSpecialMenu(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  return updateSimpleRow_(
    'Special_Menus',
    'Special_ID',
    data.specialId ||
    (data.menu || {}).id,
    data.menu || {}
  );
}


function toggleSpecialMenu(
  data
) {

  requireAdminAuth_(
    data &&
    data.adminToken
  );

  return setAvailabilityById_(
    'Special_Menus',
    'Special_ID',
    data.specialId,
    'Available',
    data.available
  );
}


function updateSimpleRow_(
  sheetName,
  idHeader,
  id,
  data
) {

  const targetId =
    String(
      id || ''
    ).trim();

  if (!targetId) {
    throw new Error(
      'ไม่พบ ID'
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      sheetName
    );

  const headers =
    getHeaders_(sheet);

  const idCol =
    headers.indexOf(
      idHeader
    );

  if (idCol < 0) {
    throw new Error(
      'ไม่พบคอลัมน์ ' +
      idHeader
    );
  }

  const values =
    sheet
      .getDataRange()
      .getValues();

  let rowNumber = -1;

  values.slice(1).forEach(
    function(row, index) {
      if (
        String(
          row[idCol] || ''
        ).trim() ===
        targetId
      ) {
        rowNumber =
          index + 2;
      }
    }
  );

  if (
    rowNumber < 0
  ) {
    throw new Error(
      'ไม่พบข้อมูล ID: ' +
      targetId
    );
  }

  const row =
    sheet
      .getRange(
        rowNumber,
        1,
        1,
        headers.length
      )
      .getValues()[0];

  headers.forEach(
    function(header, index) {

      if (
        data &&
        data[header] !== undefined
      ) {
        row[index] =
          data[header];
      }
    }
  );

  sheet
    .getRange(
      rowNumber,
      1,
      1,
      headers.length
    )
    .setValues([row]);

  return {
    success: true,
    id: targetId
  };
}


function setAvailabilityById_(
  sheetName,
  idHeader,
  id,
  availableHeader,
  available
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      sheetName
    );

  const headers =
    getHeaders_(sheet);

  const idCol =
    headers.indexOf(
      idHeader
    );

  const availableCol =
    headers.indexOf(
      availableHeader
    );

  if (
    idCol < 0 ||
    availableCol < 0
  ) {
    throw new Error(
      'ไม่พบคอลัมน์ที่ต้องการ'
    );
  }

  const values =
    sheet
      .getRange(
        2,
        idCol + 1,
        Math.max(
          sheet.getLastRow() - 1,
          0
        ),
        1
      )
      .getValues();

  let rowNumber = -1;

  values.forEach(
    function(row, index) {
      if (
        String(
          row[0] || ''
        ).trim() ===
        String(id || '').trim()
      ) {
        rowNumber =
          index + 2;
      }
    }
  );

  if (
    rowNumber < 0
  ) {
    throw new Error(
      'ไม่พบ ID: ' +
      id
    );
  }

  sheet
    .getRange(
      rowNumber,
      availableCol + 1
    )
    .setValue(
      Boolean(
        available
      )
    );

  return {
    success: true,
    id: String(id),
    available:
      Boolean(available)
  };
}


function deleteById_(
  sheetName,
  idHeader,
  id
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      sheetName
    );

  const headers =
    getHeaders_(sheet);

  const idCol =
    headers.indexOf(
      idHeader
    );

  if (idCol < 0) {
    throw new Error(
      'ไม่พบคอลัมน์ ' +
      idHeader
    );
  }

  const values =
    sheet
      .getRange(
        2,
        idCol + 1,
        Math.max(
          sheet.getLastRow() - 1,
          0
        ),
        1
      )
      .getValues();

  let rowNumber = -1;

  values.forEach(
    function(row, index) {
      if (
        String(
          row[0] || ''
        ).trim() ===
        String(id || '').trim()
      ) {
        rowNumber =
          index + 2;
      }
    }
  );

  if (
    rowNumber < 0
  ) {
    throw new Error(
      'ไม่พบ ID: ' +
      id
    );
  }

  sheet.deleteRow(
    rowNumber
  );

  return {
    success: true,
    id: String(id)
  };
}


function testDailyDishLibrary() {
  Logger.log(
    JSON.stringify(
      getDailyDishLibrary(),
      null,
      2
    )
  );
}


function testGetDailyDishLibrary() {
  Logger.log(
    JSON.stringify(
      getDailyDishLibrary(),
      null,
      2
    )
  );
}


function cleanupDuplicateDailyDishes() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Daily_Dishes'
    );

  if (!sheet) {
    throw new Error(
      'ไม่พบ Sheet Daily_Dishes'
    );
  }

  const headers =
    getHeaders_(sheet);

  const dateCol =
    headers.indexOf(
      'Dish_Date'
    );

  const nameCol =
    headers.indexOf(
      'Dish_Name'
    );

  if (
    dateCol < 0 ||
    nameCol < 0
  ) {
    throw new Error(
      'Daily_Dishes ขาด Dish_Date หรือ Dish_Name'
    );
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return {
      success: true,
      removed: 0
    };
  }

  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();

  const seen = {};
  const removeRows = [];

  rows.forEach(function(row, index) {

    const key =
      normalizeDateValue_(
        row[dateCol]
      ) +
      '|' +
      String(
        row[nameCol] || ''
      ).trim();

    if (
      seen[key]
    ) {
      removeRows.push(
        index + 2
      );
    } else {
      seen[key] = true;
    }
  });

  removeRows.reverse().forEach(
    function(rowNumber) {
      sheet.deleteRow(
        rowNumber
      );
    }
  );

  return {
    success: true,
    removed:
      removeRows.length
  };
}


function resetPickupSlots20Min() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.SHEETS.SLOTS
    );

  if (!sheet) {
    throw new Error(
      'ไม่พบ Sheet Pickup_Slots'
    );
  }

  const headers =
    getHeaders_(sheet);

  sheet.clearContents();

  sheet
    .getRange(
      1,
      1,
      1,
      5
    )
    .setValues([
      [
        'Slot_ID',
        'Start_Time',
        'End_Time',
        'Capacity',
        'Active'
      ]
    ]);

  const start =
    8 * 60;

  const end =
    17 * 60;

  const interval =
    20;

  const capacity =
    10;

  const rows = [];
  let slotNumber = 1;

  for (
    let current = start;
    current < end;
    current += interval
  ) {

    const next =
      Math.min(
        current + interval,
        end
      );

    const startHour =
      Math.floor(
        current / 60
      );

    const startMinute =
      current % 60;

    const endHour =
      Math.floor(
        next / 60
      );

    const endMinute =
      next % 60;

    const startText =
      String(
        startHour
      ).padStart(2, '0') +
      ':' +
      String(
        startMinute
      ).padStart(2, '0');

    const endText =
      String(
        endHour
      ).padStart(2, '0') +
      ':' +
      String(
        endMinute
      ).padStart(2, '0');

    rows.push([
      'S' +
      String(
        slotNumber
      ).padStart(3, '0'),
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
        5
      )
      .setValues(
        rows
      );
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    count: rows.length,
    capacity: capacity,
    firstSlot:
      rows[0][1] +
      ' - ' +
      rows[0][2],
    lastSlot:
      rows[rows.length - 1][1] +
      ' - ' +
      rows[rows.length - 1][2]
  };
}


function testCustomerDailyDishes() {

  const data =
    getInitialData();

  Logger.log(
    'APP VERSION = ' +
    data.appVersion
  );

  Logger.log(
    'DAILY DISH COUNT = ' +
    (data.dailyDishes || [])
      .length
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
    String(
      GITHUB_PROD_IMAGE_DIR_ ||
      'images/menu'
    ).replace(
      /\/$/,
      ''
    ) +
    '/test-store-' +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyyMMdd-HHmmss'
    ) +
    '-' +
    Utilities.getUuid()
      .replace(
        /-/g,
        ''
      )
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
    String(
      userId || ''
    ).trim();

  orderId =
    String(
      orderId || ''
    ).trim();

  orderDate =
    String(
      orderDate || ''
    ).trim();

  if (!userId) {
    return {
      success: false,
      sent: false,
      reason:
        'NO_LINE_USER_ID'
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

  const shopName =
    String(
      getStoreConfig_().SHOP_NAME ||
      'ร้านอาหาร'
    );

  const message =
    '✅ อาหารพร้อมรับแล้วครับ!\n\n' +
    'ร้าน: ' +
    shopName +
    '\n' +
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

  const TEST_ORDER_ID =
    'A005';

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

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

  const userIdCol = 3;

  let target = null;

  rows.forEach(function(row) {

    if (target) {
      return;
    }

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

  const orderDate =
    normalizeDateValue_(
      target[dateCol]
    );

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


function sendNewOrderGroupNotificationProd_(
  orderId,
  orderData,
  total,
  detailRows,
  orderDate,
  orderTime
) {

  const props =
    PropertiesService
      .getScriptProperties();

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
    CASH: '💵 เงินสด',
    QR: '📱 PromptPay',
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

      itemLines.push(
        line
      );
    }
  );

  if (
    itemLines.length === 0
  ) {
    itemLines.push(
      '• ดูรายละเอียดในหน้า Admin'
    );
  }

  const shopName =
    String(
      getStoreConfig_().SHOP_NAME ||
      'ร้านอาหาร'
    );

  let message =
    '🔔 ออเดอร์ใหม่!\n\n' +
    'ร้าน: ' +
    shopName +
    '\n' +
    'เลขออเดอร์: ' +
    orderId +
    '\n' +
    (
      orderTime
        ? '🕐 เวลา: ' +
          orderTime +
          '\n'
        : ''
    ) +
    '\n👤 ลูกค้า: ' +
    (customerName || '-') +
    '\n\n🍛 รายการ\n' +
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
    '\n\n👉 กรุณาเปิดหน้า Admin เพื่อรับออเดอร์';

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
            to: groupId,
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

  rows.forEach(function(row) {

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
  });

  return map;
}


function testShopGroupPushProd() {

  const props =
    PropertiesService
      .getScriptProperties();

  const groupId =
    String(
      props.getProperty(
        'SHOP_NOTIFY_GROUP_ID'
      ) || ''
    ).trim();

  if (!groupId) {
    throw new Error(
      'ไม่พบ SHOP_NOTIFY_GROUP_ID ใน Store Instance'
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
      'ไม่พบ LINE_CHANNEL_ACCESS_TOKEN ใน Store Instance'
    );
  }

  const shopName =
    String(
      getStoreConfig_().SHOP_NAME ||
      'ร้านอาหาร'
    );

  const message =
    '🔔 ทดสอบระบบ Store Instance\n\n' +
    'ร้าน: ' +
    shopName +
    '\n' +
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
            to: groupId,
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
    'STORE GROUP PUSH STATUS:',
    status
  );

  console.log(
    'STORE GROUP PUSH RESPONSE:',
    text
  );

  if (
    status < 200 ||
    status >= 300
  ) {
    throw new Error(
      'LINE Group push ไม่สำเร็จ (' +
      status +
      '): ' +
      text
    );
  }

  return {
    success: true,
    sent: true,
    httpStatus: status
  };
}
