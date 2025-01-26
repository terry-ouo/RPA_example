
import { test, expect } from '@playwright/test';
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
let excelData;

function excelToJson(filePath) {
  const workbook = xlsx.readFile(filePath); // 读取 Excel 文件
  const sheetName = workbook.SheetNames[0]; // 获取第一个工作表的名称
  const worksheet = workbook.Sheets[sheetName]; // 获取工作表
  return xlsx.utils.sheet_to_json(worksheet); // 将工作表转换为 JSON
}


test('should convert Excel to JSON correctly', () => {
  const jsonData = excelToJson('./challenge.xlsx');
  console.log(jsonData);
  excelData = jsonData;
  console.log(excelData);
  expect(jsonData).resolves;
});

test('get started link', async ({ page }) => {
  // 监听所有 HTTP 响应
  page.on('response', (response) => {
    // 如果是 404 错误，则忽略
    if (response.status() === 404) {
      console.log(`Ignored 404 error: ${response.url()}`);
    }
  });

  await page.goto('https://rpachallenge.com/');

  // 定位具有 ng-reflect-name="labelEmail" 的 input 元素
  const inputField = await page.locator('input[ng-reflect-name="labelEmail"], input[_ngcontent-c2="labelEmail"]');
  // 填写输入框
  await inputField.fill('test123@gmail.com')
  console.log('inputField.inputValue', inputField.inputValue);
  const submit = await page.locator('type="submit');
  await submit.click;
  // Expects page to have a heading with the name of Installation.
  expect(inputField.inputValue == 'test');
});
