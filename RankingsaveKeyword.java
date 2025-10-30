package SEOKeywordPresencce.Classes;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.SocketTimeoutException;
import java.net.URISyntaxException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Set;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FileUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import org.openqa.selenium.Cookie;
import com.twocaptcha.TwoCaptcha;
import com.twocaptcha.captcha.ReCaptcha;
import java.io.File;
import java.io.IOException;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FileUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import org.openqa.selenium.Cookie;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import SEOKeywordPresencce.Repository.Sendmail;
import org.apache.poi.hssf.usermodel.HSSFCell;

public class RankingsaveKeyword {

	static SEOKeywordPresencce.Driver.SingletonClassChromeNew singleton = SEOKeywordPresencce.Driver.SingletonClassChromeNew
			.getInstanceOfSingletonBrowserClass();
	static WebDriver driver = singleton.getDriver();
	static int rowcount = 0;
	static boolean error = false;
	static int passcount = 0;
	static int urlnumbercount = 0;
	static int failcount = 0;
	static int organicurlnumbercount = 0;
	static int organicpasscount = 0;
	static int organicmapurlnumbercount = 0;
	static boolean flag = false;
	static boolean organicfail = false;
	static boolean human = false;
	static WebElement element = null;
	static int count = 1;
	static int i = 0;
	static String brandName = "";
	static String sidelines = null;
	static List<String> toplist = new ArrayList();
	static String topten = null;
	static boolean toptenpresent = false;
	static String websitetext = null;
	static String googleWebLink = null;
	static String GooglelinkText = null;
	static int valueoftoptenis = 0;
	static boolean presentontopten = false;
	static boolean homepage = false;
	static int mydatacount = 0;
	static int tennumbercount = 0;
	static String pagename = null;
	static String organicstatus = null;
	static String microsite = null;
	static boolean micrositepresent = false;
	static boolean organicnearbypresence = false;
	static int countorganicnearby = 0;
	static String nearbytext = null;
	static Duration timezone = Duration.ofMillis(1000);;
	static WebDriverWait wait = new WebDriverWait(driver, timezone);
	static String finalykeyword = null;
	static String NearbyLocality = null;
	static JavascriptExecutor js = (JavascriptExecutor) driver;
	static int nearbycount = 0;
	static String mainKeyword = null;
	static int NolocalitylinkonGoogle = 0;
	static boolean nearbyPass = false;
	static String nearbybuffertext = null;
	static int rowNum = 1;
	static Workbook workbook = new XSSFWorkbook();
	static org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet("geo0");
	static List<WebElement> card2text2 = null;
	static int initialSize = 0;
	static int newSize = 0;
	static int size = 0;
	static List<WebElement> allLinks = null;
	static boolean keywordpresence = false;
	static int lastSlashIndex = 0;
	static List<WebElement> organicsection = null;
	static String organicurls = null;
	static int homepasscount = 0;
	static boolean homepagepresent = false;
	static String currenturl = null;
	static String Organic_URL_Status = null;
	static String Map_Pack_Status = null;
	static boolean ignoremap = false;

	public RankingsaveKeyword() throws IOException {
		int newrow = 0;
		org.apache.poi.ss.usermodel.Row row1 = sheet.createRow(newrow);
		ArrayList<String> names = new ArrayList<String>(Arrays.asList("Cateory", "Client StoreID",
				"Actual Client StoreID", "Business Name","State","City", "Website", "Keyword Serach", "GMB Rank", "Map Pack Status",
				"SERP Rank", "Organic URL Status", "Organic URL Page Name", "Home page presence", "Keyword Presence"));

		int c = 0;
		for (String cellName : names) {
			org.apache.poi.ss.usermodel.Cell cell = row1.createCell(c++);
			cell.setCellValue(cellName);
		}
		FileOutputStream fileOut3 = new FileOutputStream("geo0");
		workbook.write(fileOut3);

	}

	@Test(dataProvider = "dataprov")
	public static void getdata(String StoreID, String ActualStoreID, String Client_Name, String BusinessName,
			String Keyword, String Locality, String City, String State, String SILat, String SILong, String Website,
			String Domain) throws InterruptedException, IOException, ClassNotFoundException, URISyntaxException {
		toptenpresent = false;
		homepage = false;
		pagename = null;
		urlnumbercount = 0;
		organicpasscount = 0;
		tennumbercount = 0;
		flag = false;
		human = false;
		toptenpresent = false;
		organicfail = false;
		presentontopten = false;
		homepasscount = 0;
		homepagepresent = false;
		pagename = null;
		Organic_URL_Status = null;
		Map_Pack_Status = null;
		Actions actions = new Actions(driver);

		try {

			flag = false;
			toptenpresent = false;
			organicnearbypresence = false;
			mydatacount++;
			System.out.println(mydatacount);
			tennumbercount = 0;
			urlnumbercount = 0;
			size = 0;
			count = 0;
			keywordpresence = false;
			homepasscount = 0;
			homepagepresent = false;
			pagename = null;
			Organic_URL_Status = null;
			Map_Pack_Status = null;

			// create excel------------------------------------------------------
			org.apache.poi.ss.usermodel.Row row = sheet.createRow(rowNum);
			int cellUrl = 0;
			org.apache.poi.ss.usermodel.Cell cellurl = row.createCell(cellUrl);
			cellurl.setCellValue(Client_Name);
			int cellUrlone = 1;
			org.apache.poi.ss.usermodel.Cell cellurlone = row.createCell(cellUrlone);
			cellurlone.setCellValue(StoreID);
			int cellUrltwo = 2;
			org.apache.poi.ss.usermodel.Cell cellurlthree = row.createCell(cellUrltwo);
			cellurlthree.setCellValue(ActualStoreID);
			int cellUrlfour = 3;
			org.apache.poi.ss.usermodel.Cell cellurlfive = row.createCell(cellUrlfour);
			cellurlfive.setCellValue(BusinessName);			
			int cellUrlfour1 = 4;
			org.apache.poi.ss.usermodel.Cell cellurlfive1 = row.createCell(cellUrlfour1);
			cellurlfive1.setCellValue(State);			
			int cellUrlfour2 = 5;
			org.apache.poi.ss.usermodel.Cell cellurlfive2 = row.createCell(cellUrlfour2);
			cellurlfive2.setCellValue(City);
			int cellUrlsix = 6;
			org.apache.poi.ss.usermodel.Cell cellurlsix = row.createCell(cellUrlsix);
			cellurlsix.setCellValue(Website);
			FileOutputStream fileOut1 = new FileOutputStream("geo0.xlsx");
			workbook.write(fileOut1);

			String[] keywordbuffervalue = Keyword.split(",");

			// Replace with your actual data
			System.out.println("KS");
			for (String keywordValue : keywordbuffervalue) {
				countorganicnearby = 0;
				count = 0;
				flag = false;
				toptenpresent = false;
				tennumbercount = 0;
				urlnumbercount = 0;
				organicnearbypresence = false;
				initialSize = 0;
				keywordpresence = false;
				homepasscount = 0;
				homepagepresent = false;
				pagename = null;
				System.out.println("*********************************");
				finalykeyword = keywordValue + " in " + Locality + ", " + City;
				// finalykeyword = keywordValue + " in " + Locality;

				System.out.println(finalykeyword);

				if (Website != null && !Website.isEmpty()) {
//					// Load the page source as a string
//					System.out.println("Brand Website is - " + Website);
//					driver.get(Website);
//					currenturl = driver.getCurrentUrl();
//					Thread.sleep(3000);
//
//					String pageSource = driver.getPageSource();
//					System.out.println("Enter 2");
//					// Count the occurrences of the URL in the page source
//					count = countOccurrences(pageSource, keywordValue);
//
//
//					System.out.println("Number of occurrences of " + keywordValue + " in the page source: " + count);
					if (count >= 0) {
						keywordpresence = true;
					}
				}

				if (keywordpresence == true) {
					driver.get("https://www.google.com/");
					element = driver.findElement(By.name("q"));
					Thread.sleep(5000);
					
					if (driver.getCurrentUrl().contains("https://www.google.com/sorry")) {
						Thread.sleep(2000);
						System.out.println("first attempt");
						
						System.out.println("Retrieve cuurent browser 1 cookies");
					    Set<Cookie> cookies = driver.manage().getCookies();
					    for (Cookie cookie : cookies) {
					        System.out.println("Cookie: " + cookie.getName() + " = " + cookie.getValue());
					    }
						
						 String cookieString = "foo=bar; baz=1";
						 
						 
						  addCookiesFromString(driver, cookieString);
						   driver.navigate().refresh();

							System.out.println("Retrieve after existing cookies applied - 1");
						    Set<Cookie> cookies2 = driver.manage().getCookies();
						    for (Cookie cookie : cookies2) {
						        System.out.println("Cookie1: " + cookie.getName() + " = " + cookie.getValue());
						    }


						String pageSource = driver.getPageSource();
						System.out.println("page source started first attempt---");
						System.out.println(pageSource);

						// Parse the page source with Jsoup
						Document doc = Jsoup.parse(pageSource);

						String datakeypresent = null;
						String dataskey = null;
						Elements elementsWithSitekey = doc.select("[data-sitekey]");
						for (org.jsoup.nodes.Element element : elementsWithSitekey) {
							System.out.println("Data-Sitekey: " + element.attr("data-sitekey"));
							datakeypresent = element.attr("data-sitekey");
							dataskey = element.attr("data-s");
							break;

						}

						if (datakeypresent == null) {
							datakeypresent = "6LfD3PIbAAAAAJs_eEHvoOl75_83eXSqpPSRFJ_u";
						}
						System.out.println("Second attempt-" + datakeypresent);
						System.out.println("second dataskey will be-" + dataskey);

						String Capurl = driver.getCurrentUrl();
						System.out.println("CAPTCHA URL: " + Capurl);

						TwoCaptcha solver = new TwoCaptcha("ac1468bf50faa8856d504c04deda4f7e");
						ReCaptcha captcha = new ReCaptcha();
						captcha.setSiteKey(datakeypresent);
						captcha.setData(dataskey);
						captcha.setUrl(Capurl);
						captcha.setInvisible(false);
						// Configure proxy
						String proxyType = "http"; // Proxy type: "http", "socks4", or "socks5"
						String proxyAddress = "43.152.113.55"; // Replace with your proxy address
						int proxyPort = 2334; // Replace with your proxy port
						String proxyLogin = "u0c0cc529557505b5-zone-custom-region-rsa"; // Replace with your
																						// proxy
																						// username (if
																						// required)
						String proxyPassword = "u0c0cc529557505b5"; // Replace with your proxy password (if
																	// required)
						// Format the proxy string based on whether authentication is required
						String proxyDetails;
						if (proxyLogin != null && !proxyLogin.isEmpty()) {
							proxyDetails = String.format("%s:%s@%s:%d", proxyLogin, proxyPassword, proxyAddress,
									proxyPort);
						} else {
							proxyDetails = String.format("%s:%d", proxyAddress, proxyPort);
						}
						// Pass proxy type and details to setProxy
						captcha.setProxy(proxyType, proxyDetails);

						captcha.setAction("verify");
						try {
							solver.setDefaultTimeout(120);
							solver.setRecaptchaTimeout(600);
							solver.setPollingInterval(10);
							solver.solve(captcha);
							// solver.setCallback("https://2captcha.com/blog/captcha-bypass-in-selenium");
							String captchaSolution = captcha.getCode(); // Retrieve solved CAPTCHA token
							System.out.println("Captcha solved: " + captchaSolution);
							System.out.println("again attempt 3");

							JavascriptExecutor js3 = (JavascriptExecutor) driver;
							js3.executeScript(
									"document.getElementById('g-recaptcha-response').value = arguments[0];",
									captchaSolution);
							System.out.println("again attempt 4");
							// Log: Verify the new value of 'g-recaptcha-response'
							Object updatedValue = js3.executeScript(
									"return document.getElementById('g-recaptcha-response')?.value;");
							System.out.println("again attempt 5");
							System.out.println("Agin Updated value of 'g-recaptcha-response': " + updatedValue);
							WebElement form = driver.findElement(By.id("captcha-form"));
							form.submit();
							System.out.println("again attempt 6");
							// Wait for form submission response
							System.out.println("again attempt 7");
							System.out.println("Form submitted successfully!");
							Thread.sleep(1000);

						} catch (Exception e) {
							System.out.println("again Error occurred: " + e.getMessage());
							System.out.println(" again syso occurred-1: " + e.getLocalizedMessage());
							System.out.println("Print Stacks agin- " + e.getStackTrace());
							System.out.println("again attempt 8");
							// Print stack trace
							System.out.println("\nStack Trace:");
							e.printStackTrace();

							// Get stack trace as an array
							StackTraceElement[] stackTrace = e.getStackTrace();
							System.out.println("\nFormatted Stack Trace:");
							for (StackTraceElement element : stackTrace) {
								System.out.println("  at " + element);
							}

						}

					}
					Thread.sleep(2000);

					System.out.println("finalykeyword-----" + finalykeyword);
					element.sendKeys(finalykeyword);
					Thread.sleep(7000);
					element.sendKeys(Keys.ENTER);
					Thread.sleep(1000);
					if (driver.getCurrentUrl().contains("https://www.google.com/sorry")) {
						Thread.sleep(2000);
						System.out.println("first attempt");
						
						System.out.println("Retrieve cuurent browser 1 cookies");
					    Set<Cookie> cookies = driver.manage().getCookies();
					    for (Cookie cookie : cookies) {
					        System.out.println("Cookie: " + cookie.getName() + " = " + cookie.getValue());
					    }
						
						 String cookieString = "foo=bar; baz=1";
						 
						 
						  addCookiesFromString(driver, cookieString);
						   driver.navigate().refresh();

							System.out.println("Retrieve after existing cookies applied - 1");
						    Set<Cookie> cookies2 = driver.manage().getCookies();
						    for (Cookie cookie : cookies2) {
						        System.out.println("Cookie1: " + cookie.getName() + " = " + cookie.getValue());
						    }


						String pageSource = driver.getPageSource();
						System.out.println("page source started first attempt---");
						System.out.println(pageSource);

						// Parse the page source with Jsoup
						Document doc = Jsoup.parse(pageSource);

						String datakeypresent = null;
						String dataskey = null;
						Elements elementsWithSitekey = doc.select("[data-sitekey]");
						for (org.jsoup.nodes.Element element : elementsWithSitekey) {
							System.out.println("Data-Sitekey: " + element.attr("data-sitekey"));
							datakeypresent = element.attr("data-sitekey");
							dataskey = element.attr("data-s");
							break;

						}

						if (datakeypresent == null) {
							datakeypresent = "6LfD3PIbAAAAAJs_eEHvoOl75_83eXSqpPSRFJ_u";
						}
						System.out.println("Second attempt-" + datakeypresent);
						System.out.println("second dataskey will be-" + dataskey);

						String Capurl = driver.getCurrentUrl();
						System.out.println("CAPTCHA URL: " + Capurl);

						TwoCaptcha solver = new TwoCaptcha("ac1468bf50faa8856d504c04deda4f7e");
						ReCaptcha captcha = new ReCaptcha();
						captcha.setSiteKey(datakeypresent);
						captcha.setData(dataskey);
						captcha.setUrl(Capurl);
						captcha.setInvisible(false);
						// Configure proxy
						String proxyType = "http"; // Proxy type: "http", "socks4", or "socks5"
						String proxyAddress = "43.152.113.55"; // Replace with your proxy address
						int proxyPort = 2334; // Replace with your proxy port
						String proxyLogin = "u0c0cc529557505b5-zone-custom-region-rsa"; // Replace with your
																						// proxy
																						// username (if
																						// required)
						String proxyPassword = "u0c0cc529557505b5"; // Replace with your proxy password (if
																	// required)
						// Format the proxy string based on whether authentication is required
						String proxyDetails;
						if (proxyLogin != null && !proxyLogin.isEmpty()) {
							proxyDetails = String.format("%s:%s@%s:%d", proxyLogin, proxyPassword, proxyAddress,
									proxyPort);
						} else {
							proxyDetails = String.format("%s:%d", proxyAddress, proxyPort);
						}
						// Pass proxy type and details to setProxy
						captcha.setProxy(proxyType, proxyDetails);

						captcha.setAction("verify");
						try {
							solver.setDefaultTimeout(120);
							solver.setRecaptchaTimeout(600);
							solver.setPollingInterval(10);
							solver.solve(captcha);
							// solver.setCallback("https://2captcha.com/blog/captcha-bypass-in-selenium");
							String captchaSolution = captcha.getCode(); // Retrieve solved CAPTCHA token
							System.out.println("Captcha solved: " + captchaSolution);
							System.out.println("again attempt 3");

							JavascriptExecutor js3 = (JavascriptExecutor) driver;
							js3.executeScript(
									"document.getElementById('g-recaptcha-response').value = arguments[0];",
									captchaSolution);
							System.out.println("again attempt 4");
							// Log: Verify the new value of 'g-recaptcha-response'
							Object updatedValue = js3.executeScript(
									"return document.getElementById('g-recaptcha-response')?.value;");
							System.out.println("again attempt 5");
							System.out.println("Agin Updated value of 'g-recaptcha-response': " + updatedValue);
							WebElement form = driver.findElement(By.id("captcha-form"));
							form.submit();
							System.out.println("again attempt 6");
							// Wait for form submission response
							System.out.println("again attempt 7");
							System.out.println("Form submitted successfully!");
							Thread.sleep(1000);

						} catch (Exception e) {
							System.out.println("again Error occurred: " + e.getMessage());
							System.out.println(" again syso occurred-1: " + e.getLocalizedMessage());
							System.out.println("Print Stacks agin- " + e.getStackTrace());
							System.out.println("again attempt 8");
							// Print stack trace
							System.out.println("\nStack Trace:");
							e.printStackTrace();

							// Get stack trace as an array
							StackTraceElement[] stackTrace = e.getStackTrace();
							System.out.println("\nFormatted Stack Trace:");
							for (StackTraceElement element : stackTrace) {
								System.out.println("  at " + element);
							}

						}

					}
					Thread.sleep(2000);
					js.executeScript("window.scrollBy(0,100)", "");
					card2text2 = null;

					card2text2 = driver.findElements(By.cssSelector("#search .yuRUbf"));
					if (card2text2.size() > 0) {

						Thread.sleep(2000);
						System.out.println("Search 2");
						linkloop: for (WebElement optionnearby : card2text2) {
							Thread.sleep(1000);
							List<WebElement> allLinks = optionnearby.findElements(By.tagName("a"));
							for (WebElement linknearby : allLinks) {
								Thread.sleep(1000);
								System.out.println(linknearby.getAttribute("href"));
								String urlsnearby = linknearby.getAttribute("href");
								js.executeScript("window.scrollBy(0,100)", "");
								System.out.println("****************" + countorganicnearby);
								if (urlsnearby.contains(Domain)) {

									System.out.println("pass 1");
									countorganicnearby++;
									organicnearbypresence = true;

									if (urlsnearby.contains("/location/") || urlsnearby.equals(Domain)) {
										pagename = "Store Locator";
									}

									else if (urlsnearby.contains("Timeline?tag")) {
										pagename = "Timeline Tags";
									} else if (urlsnearby.contains("TimelineDetails")) {
										pagename = "Timeline Details";
									} else if (urlsnearby.contains("search=")) {
										pagename = "Store Locator";
									}

									else {
										// Find the last occurrence of '/'
										lastSlashIndex = urlsnearby.lastIndexOf('/');
										if (lastSlashIndex != -1) {
											// Remove the leading '/' and the part after the last '/'
											pagename = urlsnearby.substring(lastSlashIndex + 1);
											System.out.println("Extracted part: " + pagename);
											
											if(pagename.contains(City)||pagename.contains(State)) {
												pagename = "Store Locator";
											}
											if(pagename.contains("%2")||pagename.contains(Locality)) {
												pagename = "Store Locator";
											}
										} else {
											// Handle the case when there is no '/'
											System.out.println("No '/' found in the URL.");
										}
									}

									break linkloop;

								}

								else {
									System.out.println("fail");
									countorganicnearby++;

								}
							}

							System.out.println("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@3");

						}
					}

					// Home page
					js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
					Outerloop2: if (driver.findElements(By.cssSelector(".yuRUbf")).size() > 0) {
						organicsection = driver.findElements(By.cssSelector("#search .yuRUbf"));
						for (WebElement option : organicsection) {
							allLinks = option.findElements(By.tagName("a"));
							for (WebElement link : allLinks) {
								organicurls = link.getAttribute("href");
								System.out.println(organicurls);
								js.executeScript("window.scrollBy(0,100)", "");

								if (organicurls.endsWith("/Home") && (organicurls.startsWith(Domain))) {
									System.out.println("Here for Home");
									homepasscount++;
									homepagepresent = true;
									break Outerloop2;

								}

								else {
									System.out.println("Here for Home else");
									homepasscount++;

								}

							}
						}
					}
					js.executeScript("window.scrollTo(0, document.body.scrollHeight)");

					if (driver.findElements(By.cssSelector("h2.qrShPb.pXs6bb.PZPZlf.q8U8x.aTI8gc,.PZPZlf.ssJ7i.B5dxMb"))
							.size() > 0) {
						System.out.println("Searching on sidelines");
						sidelines = driver
								.findElement(
										By.cssSelector(".PZPZlf.ssJ7i.B5dxMb , h2.qrShPb.pXs6bb.PZPZlf.q8U8x.aTI8gc"))
								.getText();
						System.out.println("sidelines--- " + sidelines);

						if (sidelines.contains(BusinessName)) {
							tennumbercount = 1;
							toptenpresent = true;
						}
					}

					// GMB code---------------------------
					if (toptenpresent == false) {
						driver.get("https://www.google.com/maps/search/" + keywordValue + " in " + Locality + ", "
								+ City + "/@" + SILat + "," + SILong + "z?entry=ttu");
//						element = driver.findElement(By.name("q"));
//						Thread.sleep(2000);
//						element.sendKeys(finalykeyword);
//						Thread.sleep(7000);
//						element.sendKeys(Keys.ENTER);
						Thread.sleep(2000);
						JavascriptExecutor js = (JavascriptExecutor) driver;
						System.out.println("Top1");
						Thread.sleep(2500);
						List<WebElement> linksimage = driver.findElements(By.cssSelector(".hfpxzc"));
						System.out.println("Top2");
						Thread.sleep(1000);
						initialSize = linksimage.size();
						System.out.println("initialSize" + initialSize);
						newSize = 0;
						while (newSize != linksimage.size()) {
							toplist.clear(); // Clear the list before each iteration
							initialSize = linksimage.size();
							System.out.println("Top3");
							for (WebElement imagelink : linksimage) {
								Thread.sleep(3000);
								System.out.println("Top4");
								actions.moveToElement(imagelink);
								js.executeScript("arguments[0].scrollIntoView()", imagelink);
								Thread.sleep(1000);
							}
							System.out.println("Top5");
							linksimage = driver.findElements(By.cssSelector(".hfpxzc"));
							newSize = linksimage.size();
							System.out.println(newSize);
						}
						System.out.println("Top6");
						System.out.println("Here toprank 1**********");
						size = Math.min(linksimage.size(), 20);
						outerlooptop: for (int i = 0; i < size; i++) {
							System.out.println("Top7");
							WebElement imagelink = linksimage.get(i);
							System.out.println("*******************");
							System.out.println("Here 2*******************");
							Thread.sleep(3000);
							actions.moveToElement(imagelink);
							js.executeScript("arguments[0].scrollIntoView()", imagelink);
							Thread.sleep(1000);
							websitetext = imagelink.getAttribute("aria-label");
							System.out.println("----------------" + websitetext);
							System.out.println("----------------Businessname" + BusinessName);

							if (websitetext.contains(BusinessName)) {
								Thread.sleep(5000);
								System.out.println("Enter in top ten");
								tennumbercount++;
								System.out.println("websitetext matcheddd");
								js.executeScript("window.scrollBy(0,100)", "");
								toptenpresent = true;
								System.out.println("value of top ten" + tennumbercount);
								break outerlooptop;
							} else {
								System.out.println("value of top ten" + tennumbercount);
								tennumbercount++;
							}
						}
					}

					if (driver.findElements(By.cssSelector("h1.DUwDvf.lfPIob")).size() > 0) {
						String BusinessnameTop = driver.findElement(By.cssSelector("h1.DUwDvf.lfPIob")).getText();
						System.out.println(BusinessnameTop);
						if (BusinessnameTop.equalsIgnoreCase(BusinessName)) {
							Thread.sleep(5000);

							tennumbercount = 1;
							System.out.println("In Map Place");
							toptenpresent = true;
						}
					}

					System.out.println("***************************************##############################8");
					if (toptenpresent == false) {
						tennumbercount = 0;
						Map_Pack_Status = "Not Present";
					}

					else {
						Map_Pack_Status = "Present";

					}
					if (organicnearbypresence == false) {
						countorganicnearby = 0;
						Organic_URL_Status = "Not Present";

					} else {
						Organic_URL_Status = "Present";
					}
					if (homepagepresent == false) {
						homepasscount = 0;

					}

					System.out.println("Top Ten value is" + tennumbercount);
					System.out.println("Value of roiw -" + rowNum);
					row = sheet.getRow(rowNum);
					if (row == null) {
						row = sheet.createRow(rowNum);
					}
					int cellNum = 7;
					int cellnum2 = 8;
					int cellnum3 = 9;
					int cellnum4 = 10;
					int cellnum5 = 11;
					int cellnum6 = 12;
					int cellnum7 = 13;
					int cellnum8 = 14;

					org.apache.poi.ss.usermodel.Cell cell = row.createCell(cellNum);
					org.apache.poi.ss.usermodel.Cell cellcountGMB = row.createCell(cellnum2);
					org.apache.poi.ss.usermodel.Cell cellcountGMBStatus = row.createCell(cellnum3);
					org.apache.poi.ss.usermodel.Cell cellcount = row.createCell(cellnum4);
					org.apache.poi.ss.usermodel.Cell cellcountstatus = row.createCell(cellnum5);
					org.apache.poi.ss.usermodel.Cell cellpagename = row.createCell(cellnum6);
					org.apache.poi.ss.usermodel.Cell cellhomecount = row.createCell(cellnum7);
					org.apache.poi.ss.usermodel.Cell cellkeywordcount = row.createCell(cellnum8);

					cell.setCellValue(finalykeyword);
					cellcountGMB.setCellValue(tennumbercount);
					cellcountGMBStatus.setCellValue(Map_Pack_Status);
					cellcount.setCellValue(countorganicnearby);
					cellcountstatus.setCellValue(Organic_URL_Status);
					cellpagename.setCellValue(pagename);
					cellhomecount.setCellValue(homepasscount);
					cellkeywordcount
							.setCellValue("Number of occurrences of '" + keywordValue + "' in the microsite: " + count);

					fileOut1 = new FileOutputStream("geo0.xlsx");
					workbook.write(fileOut1);
					rowNum++;

					System.out.println("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@5");

				}

				else {
					CellStyle orangeTextCellStyle = workbook.createCellStyle();
					orangeTextCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
					orangeTextCellStyle.setFillPattern(FillPatternType.NO_FILL);
					orangeTextCellStyle.setFont(createFontWithColor(IndexedColors.RED.getIndex()));
					System.out.println("Top Ten value is" + tennumbercount);
					System.out.println("Value of roiw -" + rowNum);
					row = sheet.getRow(rowNum);
					if (row == null) {
						row = sheet.createRow(rowNum);
					}

					int cellNum = 7;
					org.apache.poi.ss.usermodel.Cell cell = row.createCell(cellNum);
					cell.setCellValue(finalykeyword);

					int cellnum2 = 8;
					org.apache.poi.ss.usermodel.Cell cellcount = row.createCell(cellnum2);
					cellcount.setCellValue("");

					// Clear other cells for this row

					int cellnum8 = 14;

					org.apache.poi.ss.usermodel.Cell cellkeywordcount = row.createCell(cellnum8);
					cellkeywordcount.setCellValue("The report cannot be generated as the '" + keywordValue
							+ "' keyword does not exist on the microsite.");
					cellkeywordcount.setCellStyle(orangeTextCellStyle);

					fileOut1 = new FileOutputStream("geo0.xlsx");
					workbook.write(fileOut1);

					rowNum++;

				}

				System.out.println("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@6");

			}

		} catch (Exception e) {
			System.out.println("Exception occurred: " + e.getMessage());

			urlnumbercount++;
			failcount++;
			// Add conditions based on e.getMessage()
			if (e.getMessage().contains("specific error message")) {
				// Handle the specific error condition
			} else if (e.getMessage().contains("Timed out receiving message from renderer")) {
				// Handle the timeout from renderer
				System.out.println("Timed out receiving message from renderer. Taking corrective action...");
				driver.navigate().refresh();

				// Optionally, you can add additional actions or retries here

			}

			else if (e.getMessage().contains("disconnected: received Inspector.detached event")) {
				// Handle the disconnected event
				System.out.println("Disconnected: received Inspector.detached event. Taking corrective action...");
				driver.navigate().refresh();
				// Optionally, you can add additional actions or retries here

			}

			else {
				System.out.println("Catcchhh");
			}
		}

	}

	// Method to create a font with the specified color
	private static Font createFontWithColor(short color) {
		Font font = workbook.createFont();
		font.setColor(color);
		return font;
	}

	@DataProvider
	public Object[][] dataprov() throws IOException {
		System.out.println("@DataProvider");
		String[][] data = readXLSXFileurl();
		return (data);
	}

	public static String[][] readXLSXFileurl() throws IOException {
		DataFormatter formatter = new DataFormatter();
		InputStream file = new FileInputStream(
				System.getProperty("user.dir") + "\\src\\main\\java\\SEOKeywordPresencce\\Excel\\Geojit.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(file); // get my workbook
		XSSFSheet worksheet = wb.getSheetAt(0);//get my sheet from workbook
		XSSFRow Row = worksheet.getRow(0); // get my Row which start from 0

		int RowNum = worksheet.getPhysicalNumberOfRows();// count my number of Rows
		int ColNum = Row.getLastCellNum(); // get last ColNum
		int rowIndex = 0;

		String Data[][] = new String[RowNum - 1][ColNum]; // pass my count data in array

		for (int i = 0; i < RowNum - 1; i++) // Loop work for Rows
		{
			System.out.println("1");
			XSSFRow row = worksheet.getRow(i + 1);

			for (int j = 0; j < ColNum; j++) // Loop work for colNum
			{
				// System.out.println("2");
				if (row == null) {
					// System.out.println("3");
					Data[i][j] = "";
				} else {
					XSSFCell cell = row.getCell(j);
					if (cell == null) {
						// System.out.println("4");
						Data[i][j] = ""; // if it get Null value it pass no data
					} else if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
						// System.out.println("String value");
						String value = formatter.formatCellValue(cell);
						Data[i][j] = value;
					} else {
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							// Check if the cell contains a numeric value
							String value = new java.text.DecimalFormat("0").format(cell.getNumericCellValue());
							System.out.println(value);
							Data[i][j] = value;
						} else {
							// Handle other cell types (non-numeric)
							System.out.println("Cell does not contain a numeric value");
							Data[i][j] = ""; // You may want to provide a default value or handle non-numeric values
												// differently
						}
					}

				}
			}
			rowIndex++;
		}
		return Data;
	}

	public static int countOccurrences(String haystack, String needle) {
		int count = 0;
		int index = 0;
		haystack = haystack.toLowerCase();
		needle = needle.toLowerCase();

		while ((index = haystack.indexOf(needle, index)) != -1) {
			count++;
			index += needle.length();

		}

		return count;
	}
	
	public static void addCookiesFromString(WebDriver driver, String cookieString) {
		System.out.println("Added cookie method");
	    String[] cookies = cookieString.split("; "); // Split by "; " to get individual cookies
	    for (String cookie : cookies) {
	    	System.out.println("A");
	        String[] keyValue = cookie.split("=", 2); // Split into key and value
	        if (keyValue.length == 2) {
	            String name = keyValue[0].trim();
	            String value = keyValue[1].trim();
	            Cookie seleniumCookie = new Cookie(name, value);
	            driver.manage().addCookie(seleniumCookie); // Add cookie to the browser
	            System.out.println("Added cookie: " + seleniumCookie);
	        }
	    }
	}

	@AfterTest
	public void aftertest() {

		{
			Sendmail.mail("Automation: Nearby Locality Presence Report of " + brandName, brandName);

			File deletefileallure = new File(System.getProperty("user.dir") + "\\allure-results");
			for (File file : deletefileallure.listFiles()) {
				System.out.println("enter in allure");
				if (!file.isDirectory()) {
					file.getPath();
					file.delete();
					System.out.println("Files deleted" + file.getName());
				}
			}
		}
		driver.quit();
	}

}
