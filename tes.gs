// Test 1: init data
function test_stpInit() {
  var r = stpGetInitData();
  Logger.log('items: ' + r.items.length + ', customers: ' + r.customers.length);
}

// Test 2: create dummy STP
function test_stpSave() {
  var r = stpSaveNew({
    header: {
      cust: 'PT ABC',
      periode: '2026-07',
      schedule_date: '2026-07-15',
      owner_used: 'FC',
      priority: 'Normal'
    },
    items: [
      { item_code: 'PASTE_ITEM_CODE_ASLI_DISINI', qty_req: 1000, note: 'test 1' },
      { item_code: 'PASTE_ITEM_CODE_ASLI_DISINI_2', qty_req: 500 }
    ]
  });
  Logger.log(JSON.stringify(r));
}

// Test 3: list
function test_stpList() {
  var r = stpGetList({ status: 'ALL' });
  Logger.log('groups: ' + (r.groups || []).length);
}