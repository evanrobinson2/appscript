
function testCreateLineItems() {
  const oppId = '0068b00000ABCDEFG'; // example
  const newItems = [
    {
      Product__c: '01t8b00000ABCD1234',
      Price_Book__c: '01s8b00000XYZ9876',
      Quantity__c: 10,
      List_Price__c: 500.00,
      Sales_Price__c: 450.00,
      Active__c: true,
      Start_Date__c: '2025-02-01',
      End_Date__c: '2025-12-31',
      Version_Number__c: 1
    },
    {
      Product__c: '01t8b00000ABCD5678',
      Price_Book__c: '01s8b00000XYZ6543',
      Quantity__c: 3,
      List_Price__c: 1000.00,
      Sales_Price__c: 950.00,
      Active__c: true
    }
  ];

  const results = createLineItems(oppId, newItems);
  Logger.log(JSON.stringify(results, null, 2));
}
