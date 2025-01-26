import com.microsoft.playwright.*;
import java.awt.*;
import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;

public class PlaywrightDownloadFile {

    public static void main(String[] args) {
        // 初始化 Playwright
        try (Playwright playwright = Playwright.create()) {
            Browser browser = playwright.chromium().launch(new BrowserType.LaunchOptions().setHeadless(false));
            BrowserContext context = browser.newContext(new Browser.NewContextOptions().setAcceptDownloads(true));
            Page page = context.newPage();

            // 监听下载事件
            page.onDownload(download -> {
                try {
                    // 等待下载完成并保存到指定路径
                    Path downloadPath = download.path();
                    System.out.println("Downloaded file path: " + downloadPath);

                    // 打开下载的文件
                    openFile(downloadPath.toFile());
                } catch (Exception e) {
                    e.printStackTrace();
                }
            });

            // 访问网页并触发下载
            page.navigate("https://www.example.com/");
            page.click("text=Download"); // 假设页面上有一个按钮可以触发下载

            // 等待几秒，确保下载完成
            page.waitForTimeout(5000);

            // 关闭浏览器
            browser.close();
        }
    }

    // 使用系统默认应用程序打开文件
    private static void openFile(File file) {
        if (Desktop.isDesktopSupported()) {
            try {
                Desktop desktop = Desktop.getDesktop();
                desktop.open(file); // 打开文件
            } catch (Exception e) {
                System.err.println("Failed to open file: " + e.getMessage());
            }
        } else {
            System.err.println("Desktop operations are not supported on this system.");
        }
    }
}
