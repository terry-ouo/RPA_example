package com.example.RPA;

import com.microsoft.playwright.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.util.*;

public class RPA {
    private static String downloadSpace;

    public static List<Map<String, String>> readExcel(Page page) throws IOException {
        // 設定下載行為的監聽
        page.onDownload(download -> {
            try {
                Path downloadPath = download.path();
                downloadSpace = downloadPath.toString();
                System.out.println("Downloaded file path: " + downloadPath);
            } catch (Exception e) {
                e.printStackTrace();
            }
        });

        // 開啟畫面並且下載資料
        page.navigate("https://rpachallenge.com/");
        page.getByText("Download Excel").click();

        // 確保下載完成
        page.waitForTimeout(5000);

        // 讀取資料位置的資料
        FileInputStream file = new FileInputStream(downloadSpace);
        Workbook workbook = new XSSFWorkbook(file);
        //讀取工作表單(index)
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        List<Map<String, String>> data = new ArrayList<>();

        // 新增每一列
        Row headerRow = rowIterator.next();
        List<String> headers = new ArrayList<>();
        for (Cell cell : headerRow) {
            if (!cell.getStringCellValue().isEmpty()) {
                headers.add(cell.getStringCellValue());
            }
        }

        // 新增每一列的資料
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Map<String, String> rowData = new HashMap<>();
            boolean hasData = false;

            for (int i = 0; i < headers.size(); i++) {
                Cell cell = row.getCell(i);
                String value = getCellValue(cell);
                if (value != null && !value.isEmpty()) {
                    rowData.put(headers.get(i), value);
                    hasData = true;
                }
            }

            if (hasData) {
                data.add(rowData);
            }
        }

        workbook.close();
        file.close();
        return data;
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> BigDecimal.valueOf(cell.getNumericCellValue()).toString();
            default -> "";
        };
    }

    private static boolean fillData(Page page, List<Map<String, String>> data) {
        // 點擊開始
        page.locator("button.uiColorButton").click();

        for (Map<String, String> map : data) {
            page.locator("input[ng-reflect-name='labelCompanyName']").fill(map.get("Company Name"));
            page.locator("input[ng-reflect-name='labelLastName']").fill(map.get("Last Name "));
            page.locator("input[ng-reflect-name='labelEmail']").fill(map.get("Email"));
            page.locator("input[ng-reflect-name='labelAddress']").fill(map.get("Address"));
            page.locator("input[ng-reflect-name='labelFirstName']").fill(map.get("First Name"));
            page.locator("input[ng-reflect-name='labelPhone']").fill(map.get("Phone Number"));
            page.locator("input[ng-reflect-name='labelRole']").fill(map.get("Role in Company"));

            page.locator("input.uiColorButton").click();
        }

        page.waitForTimeout(5000);
        return page.getByText("Congratulations!").isVisible();
    }

    public static void main(String[] args) throws Exception {
        List<Map<String, String>> excelData;

        try (Playwright playwright = Playwright.create()) {
            Browser browser = playwright.chromium().launch(new BrowserType.LaunchOptions().setHeadless(false));
            BrowserContext context = browser.newContext();
            Page page = context.newPage();

            // 從Excel讀取資料
            excelData = readExcel(page);

            boolean result = fillData(page, excelData);

            if (result) {
                System.out.println("Execution completed successfully.");
            } else {
                System.out.println("An issue occurred during execution.");
            }

            browser.close();
        }
    }
}
