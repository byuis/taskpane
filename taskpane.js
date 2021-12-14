
console.log("starting--------------------------")


async function test(excel) {
  try {
    
      const range = excel.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "yellow";
      await excel.sync();
      console.log(`The range address was ${range.address}.`);
  } catch (error) {
    console.error(error);
  }
}
