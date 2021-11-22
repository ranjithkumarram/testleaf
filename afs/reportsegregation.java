package com.fss.afs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.TreeMap;
 
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.examples.html.HSSFHtmlHelper;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 

 
public class reportsegregation {
    static Properties prop, prop1;
    static FileInputStream fin, fin1;
    static FileOutputStream fo, fo1;
    static HashMap<String, ArrayList<String>> ggIssue = new HashMap();
 
    public static void main(String[] args) {
 
        try {
            transCount();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
 
            try {
                prop.store(fo, null);
                fo.close();
                fin.close();
 
                /*
                 * prop1.store(fo1,null); fo1.close(); fin1.close();
                 */
            } catch (IOException e1) {
                // TODO Auto-generated catch block
                e1.printStackTrace();
            }
 
        }
    }
 
    public static void initialize() {
        try {
            prop = new Properties();
            fin = new FileInputStream("D:\\OnlineTransactionsMonitor\\CommonParam.properties");
            prop.load(fin);
 
            fo = new FileOutputStream("D:\\OnlineTransactionsMonitor\\CommonParam.properties");
 
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
 
    public void propertyStore() {
        try {
            prop.store(fo, null);
            fo.close();
            fin.close();
 
            /*
             * prop1.store(fo1,null); fo1.close(); fin1.close();
             */
        } catch (IOException e1) {
            // TODO Auto-generated catch block
            e1.printStackTrace();
        }
    }
 
    // Get The Current Oman Time
    private static String getCurrentOmanTime() {
 
        // Set Oman Time zone
        java.util.TimeZone tz1 = java.util.TimeZone.getTimeZone("GMT+4");
        java.util.Calendar c1 = java.util.Calendar.getInstance(tz1);
 
        System.out.println(
                "Oman : " + c1.get(java.util.Calendar.DAY_OF_MONTH) + "\t" + c1.get(java.util.Calendar.HOUR_OF_DAY)
                        + ":" + c1.get(java.util.Calendar.MINUTE) + ":" + c1.get(java.util.Calendar.SECOND));
 
        System.out.println("Int Date:" + c1.get(java.util.Calendar.DATE));
 
        // Get Current Day as a number
        String currOmanTime = c1.get(java.util.Calendar.HOUR_OF_DAY) + ":" + c1.get(java.util.Calendar.MINUTE) + ":"
                + c1.get(java.util.Calendar.SECOND);
 
        return currOmanTime;
    }
 
    // Get The Current Bahrain Time
    private static String getCurrentBahrainTime() {
 
        // Set Bahrain Time zone
        java.util.TimeZone tz1 = java.util.TimeZone.getTimeZone("GMT+3");
        java.util.Calendar c1 = java.util.Calendar.getInstance(tz1);
 
        System.out.println(
                "Bah : " + c1.get(java.util.Calendar.DAY_OF_MONTH) + "\t" + c1.get(java.util.Calendar.HOUR_OF_DAY) + ":"
                        + c1.get(java.util.Calendar.MINUTE) + ":" + c1.get(java.util.Calendar.SECOND));
 
        // Get Current Time as a number
        String currBahrainTime = c1.get(java.util.Calendar.HOUR_OF_DAY) + ":" + c1.get(java.util.Calendar.MINUTE) + ":"
                + c1.get(java.util.Calendar.SECOND);
 
        return currBahrainTime;
    }
 
    public static void transCount() throws IOException {
 
        String timeStamp = new SimpleDateFormat("ddMM_HHmm").format(Calendar.getInstance().getTime());
        String timeStamp1 = new SimpleDateFormat("ddMM").format(Calendar.getInstance().getTime());
 
        System.out.println("Timestamp 1 :- " + timeStamp1);
 
        initialize();
 
        // Critical RC Null chcking flag
        int tester = 0;
 
        StringBuilder text = new StringBuilder();
 
        StringBuilder rcText = new StringBuilder();
 
        TreeMap<String, Integer> count = new TreeMap<>();
        TreeMap<String, Integer> visa = new TreeMap<>();
        TreeMap<String, Integer> mastercard = new TreeMap<>();
        TreeMap<String, Integer> omannet = new TreeMap<>();
        TreeMap<String, Integer> visionPlus = new TreeMap<>();
        TreeMap<String, Integer> amex = new TreeMap<>();
        TreeMap<String, Integer> other = new TreeMap<>();
        int tot;
        int bahrainInitial = 0, soharInitial = 0, AUBInitial = 0;
        int[] total = new int[6];
 
        /* Compare RC chart preparation */
 
        rcText.append("<table border='1' bordercolor='BLACK' style='border-collapse:collapse; font-family:Calibri;'>");
        rcText.append("<tr  border = '1 px solid black'>");
 
        rcText.append("<td colspan='1'  border = '1 px solid black'style ='text-align : center'>");
        rcText.append("<b>Response Code</b>" + "</td>");
        rcText.append("<td colspan='1'  border = '1 px solid black'style ='text-align : center'>");
        rcText.append("<b>Previous Count</b>" + "</td>");
        rcText.append("<td colspan='1'  border = '1 px solid black'style ='text-align : center'>");
        rcText.append("<b>Current Count</b>" + "</td>");
 
        rcText.append("</tr>");
 
        /* End of this section */
 
        HSSFWorkbook book = new HSSFWorkbook();
        HSSFSheet sheet = book.createSheet("transactions");
 
        HSSFCellStyle style = book.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont font = book.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        font.setFontName("Calibri");
        style.setFont(font);
        HSSFCellStyle style1 = book.createCellStyle();
        style1.setAlignment(HorizontalAlignment.LEFT);
        HSSFFont font1 = book.createFont();
        font1.setFontHeightInPoints((short) 12);
        font1.setBold(true);
        font1.setFontName("Calibri");
        style1.setFont(font1);
        HSSFCellStyle style2 = book.createCellStyle();
        HSSFFont font2 = book.createFont();
        font2.setFontHeightInPoints((short) 12);
        // font2.setBoldweight((short) 1000000);
        font2.setFontName("Calibri");
        style2.setAlignment(HorizontalAlignment.LEFT);
        style2.setFont(font2);
 
        // *******************************************************************************************
        try {
 
            if (new File(prop.getProperty("ManualomanInFile") + ".xlsx").exists()) {
                FileInputStream updated = new FileInputStream(prop.getProperty("ManualomanInFile") + ".xlsx");
                XSSFWorkbook workbook = new XSSFWorkbook(updated);
                // FileOutputStream fos = new FileOutputStream(new
                // File(prop.getProperty("omanBackup")));
                // workbook.write(fos);
                XSSFSheet sheeet = workbook.getSheet(workbook.getSheetName(0));
 
                rcText.append("<tr  border = '1 px solid black'>");
                rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                rcText.append("<b>========= ABO ===========</b>" + "</td></tr>");
 
                for (int i = 8; i <= sheeet.getLastRowNum() - 2; i++) {
                    XSSFRow row = sheeet.getRow(i);
                    XSSFCell cell = row.getCell(18); // Resp Code
                    String code = cell.getStringCellValue();
                    XSSFCell cell2 = row.getCell(17); // Interchange
                    String interchange = cell2.getStringCellValue();
                    cell2 = row.getCell(19); // Settlement stat ; 20 = Tran stat
                    String status = cell2.getStringCellValue();
 
                    if ((code.equals("") || code.equals(null) || code.equals(" ") || code.equals("0"))) {
                        if (row.getCell(20).getStringCellValue().equalsIgnoreCase("In progress")) {
                            code = new String(row.getCell(20).getStringCellValue());
                        } else {
                            code = row.getCell(20).getStringCellValue();
                        }
                    } else if (row.getCell(20).getStringCellValue().equalsIgnoreCase("Timeout")) {
 
                        code = code + " " + new String(row.getCell(20).getStringCellValue());
 
                    } else if (!(status.equals("Not initiated") || status.equals("Settled"))) {
                        code = code + " " + status;
                    }
 
                    /*
                     * Makarand added code if
                     * (row.getCell(19).getStringCellValue().equalsIgnoreCase(
                     * "Timeout")) { code = code + " " + new
                     * String(row.getCell(19).getStringCellValue()); }
                     */
 
                    switch (interchange) {
                    case "OMANNET":
                        total[2]++;
                        if (omannet.containsKey(code))
                            omannet.put(code, omannet.get(code) + 1);
                        else
                            omannet.put(code, 1);
                        break;
                    case "MASTERCARD":
                        total[1]++;
                        if (mastercard.containsKey(code))
                            mastercard.put(code, mastercard.get(code) + 1);
                        else
                            mastercard.put(code, 1);
                        break;
                    case "VISAAHB":
                        total[0]++;
                        if (visa.containsKey(code))
                            visa.put(code, visa.get(code) + 1);
                        else
                            visa.put(code, 1);
                        break;
                    case "VISIONPLUSAHB":
                        total[3]++;
                        if (visionPlus.containsKey(code))
                            visionPlus.put(code, visionPlus.get(code) + 1);
                        else
                            visionPlus.put(code, 1);
                        break;
                    case "AMEX":
                        total[4]++;
                        if (amex.containsKey(code))
                            amex.put(code, amex.get(code) + 1);
                        else
                            amex.put(code, 1);
                        break;
                    default:
                        if (other.containsKey(code))
                            other.put(code, other.get(code) + 1);
                        else
                            other.put(code, 1);
                    }
                    total[5]++;
                    if (count.containsKey(code)) {
                        count.put(code, count.get(code) + 1);
                    } else {
                        count.put(code, 1);
                    }
                }
                System.out.println(count);
                int success;
                if (count.get("000") != null && count.get("00") != null) {
                    success = count.get("00") + count.get("000");
                    count.put("00", success);
                    count.remove("000");
                } else if (count.get("000") == null && count.get("00") == null) {
 
                } else {
                    if (count.get("00") != null) {
                        success = count.get("00");
                    } else
                        success = count.get("000");
                    count.put("00", success);
                    count.remove("000");
                }
 
                int dispute = 0;
                if (count.get("000 Dispute Server") != null && count.get("00 Dispute Server") != null) {
                    dispute = count.get("00 Dispute Server") + count.get("000 Dispute Server");
                    count.put("00 Dispute Server", dispute);
                    count.remove("000 Dispute Server");
                } else if (count.get("000 Dispute Server") == null && count.get("00 Dispute Server") == null) {
 
                } else {
                    if (count.get("00 Dispute Server") != null) {
                        dispute = count.get("00 Dispute Server");
                    } else if (count.get("000 Dispute Server") != null)
                        dispute = count.get("000 Dispute Server");
                    count.put("00 Dispute Server", dispute);
                    count.remove("000 Dispute Server");
                }
                HSSFRow headerRow = sheet.createRow(0);
                HSSFCell headerCell = headerRow.createCell(0);
                headerCell.setCellValue("Transactions Detail report - Oman");
 
                text.append(
                        "<table border='1' bordercolor='BLACK' style='border-collapse:collapse; font-family:Calibri;'>");
                text.append("<tr  border = '1 px solid black'>");
                text.append("<td colspan='8'  border = '1 px solid black'style ='text-align : center'>");
                text.append("<b>Transactions Detail report - Oman</b>");
                text.append("</td>");
                text.append("</tr>");
                sheet.addMergedRegion(new CellRangeAddress(0, 1, 0,7));
                headerCell.setCellStyle(style);
                sheet.autoSizeColumn(0);
                headerRow = sheet.createRow(sheet.getLastRowNum() + 1);
 
                // Makarand code for new output Sheet
                String bankABO = "Financial Institution ID : AHB - Ahli Bank of Oman";
                String tranDateABO = "Transaction Date : " + getCurrentOmanTime(); // append
                                                                                    // Oman
                                                                                    // Time
                String runDateABO = "Run Date/Time : " + java.time.LocalDate.now() + " " + getCurrentOmanTime()
                        + " Asia/Muscat"; // Append IST Date and Oman Time
 
                for (int i = 1; i < 4; i++) {
                    headerRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    HSSFCell bankCell = headerRow.createCell(0);
                    sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 0, 7));
                    bankCell.setCellStyle(style1);
                    if (i == 1) {
 
                        bankCell.setCellValue(bankABO);
                        text.append("<tr>");
                        text.append("<td colspan='8'   border = '1 px solid black'>");
                        text.append("<b>" + bankABO + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                    if (i == 2) {
 
                        bankCell.setCellValue(tranDateABO);
                        text.append("<tr>");
                        text.append("<td colspan='8'   border = '1 px solid black'>");
                        text.append("<b>" + tranDateABO + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                    if (i == 3) {
 
                        bankCell.setCellValue(runDateABO);
                        text.append("<tr>");
                        text.append("<td colspan='8'   border = '1 px solid black'>");
                        text.append("<b>" + runDateABO + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                }
 
                /*
                 * for (int i = 2; i <= 5; i++) { text.append("<tr>");
                 * text.append("<td colspan='7'   border = '1 px solid black'>"
                 * ); headerRow = sheet.createRow(sheet.getLastRowNum() + 1);
                 * HSSFCell bankCell = headerRow.createCell(0);
                 * bankCell.setCellStyle(style1);
                 * bankCell.setCellValue(sheeet.getRow(i).getCell(0).toString())
                 * ; text.append("<b>" + sheeet.getRow(i).getCell(0).toString()
                 * + "</b>"); sheet.addMergedRegion(new CellRangeAddress(i, i,
                 * 0, 6)); text.append("</td>"); text.append("</tr>"); }
                 */
 
                HSSFRow rowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                text.append("<tr>");
                HSSFCell cellTotal = rowTotal.createCell(0);
 
                cellTotal.setCellStyle(style);
                sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 1));
                cellTotal.setCellValue("Total Transactions");
                text.append("<td border = '1 px solid black' colspan='2' rowspan='2' style ='text-align : center'>");
                text.append("<b>Total Transactions</b>");
                text.append("</td>");
                cellTotal = rowTotal.createCell(2);
                cellTotal.setCellStyle(style);
                cellTotal.setCellValue("VISA");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISA</b>");
                text.append("</td>");
                cellTotal = rowTotal.createCell(3);
                cellTotal.setCellStyle(style);
                cellTotal.setCellValue("MASTER CARD");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Master Card</b>");
                text.append("</td>");
                cellTotal = rowTotal.createCell(4);
                cellTotal.setCellStyle(style);
                cellTotal.setCellValue("OMANNET");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>OMANNET</b>");
                text.append("</td>");
                cellTotal = rowTotal.createCell(5);
                cellTotal.setCellStyle(style);
                cellTotal.setCellValue("VISION PLUS");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISION PLUS</b>");
                text.append("</td>");
                cellTotal = rowTotal.createCell(6);
                cellTotal.setCellStyle(style);
                cellTotal.setCellValue("AMEX");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>AMEX</b>");
                text.append("</td>");
                cellTotal = rowTotal.createCell(7);
                cellTotal.setCellStyle(style);
                cellTotal.setCellValue("GRAND TOTAL");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>GRAND TOTAL</b>");
                text.append("</td>");
                text.append("</tr>");
                rowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                text.append("<tr>");
 
                for (int i = 0; i < total.length; i++) {
                    int j = i + 2;
                    cellTotal = rowTotal.createCell(j);
                    cellTotal.setCellStyle(style2);
                    cellTotal.setCellValue(total[i]);
                    text.append("<td  border = '1 px solid black'>");
                    text.append(total[i]);
                    text.append("</td>");
                }
 
                text.append("</tr>");
                HSSFRow rowBlank = sheet.createRow(sheet.getLastRowNum() + 1);
                text.append("<tr>");
 
                HSSFCell cellHeader = rowBlank.createCell(0);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("Response code");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Response code</b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(1);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("Response code description");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Response code description/<b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(2);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("VISA");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISA</b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(3);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("MASTERCARD");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>MASTERCARD</b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(4);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("OMANNET");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>OMANNET</b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(5);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("VISIONPLUS");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISIONPLUS</b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(6);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("AMEX");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>AMEX</b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(7);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("Total count");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Total count</b>");
                text.append("</td>");
                text.append("</tr>");
                cellHeader.setCellStyle(style);
               
                sheet.autoSizeColumn(0);
                sheet.autoSizeColumn(1);
                sheet.autoSizeColumn(2);
                sheet.autoSizeColumn(3);
                sheet.autoSizeColumn(4);
                sheet.autoSizeColumn(5);
                sheet.autoSizeColumn(6);
                sheet.autoSizeColumn(7);
 
                System.out.println("Oman Count : " + count);
 
                int length = sheet.getLastRowNum() + 1;
                Set<String> keys = count.keySet();
                // System.out.println(keys);
                Iterator<String> it = keys.iterator();
                for (int i = 0; i < count.size(); i++) {
                    text.append("<tr>");
                    String temp = it.next();
                    HSSFRow dataRow = sheet.createRow(length++);
                    dataRow.setRowStyle(style);
                    HSSFCell dataCell = dataRow.createCell(0);
                    dataCell.setCellStyle(style2);
                    dataCell.setCellValue(temp);
                    text.append("<td border = '1 px solid black'>");
                    text.append(temp);
                    text.append("</td>");
                    if (temp.length() > 4) {
                    }
                    dataCell = dataRow.createCell(1);
                    dataCell.setCellStyle(style2);
                    // if (flag) {
                    // dataCell.setCellValue(settlementStatus);
                    //
                    // } else {
                    if (prop.getProperty(temp) != null) {
                        dataCell.setCellValue(prop.getProperty(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append(prop.getProperty(temp));
                        text.append("</td>");
                    }
 
                    else {
                        dataCell.setCellValue("unknown response code/ unknown status");
                        text.append("<td border = '1 px solid black'>");
                        text.append("unknown response code/ unknown status");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(2);
                    dataCell.setCellStyle(style2);
                    if (visa.get(temp) != null) {
                        text.append("<td border = '1 px solid black'>");
                        text.append((visa.get(temp)));
                        text.append("</td>");
                        dataCell.setCellValue(visa.get(temp));
                    }
 
                    else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
 
                    dataCell = dataRow.createCell(3);
                    dataCell.setCellStyle(style2);
                    if (mastercard.get(temp) != null) {
                        text.append("<td border = '1 px solid black'>");
                        text.append((mastercard.get(temp)));
                        text.append("</td>");
                        dataCell.setCellValue(mastercard.get(temp));
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                   
                    dataCell = dataRow.createCell(4);
                    dataCell.setCellStyle(style2);
                    if (temp.equals("00")) {
                        temp = "000";
                        if (omannet.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((omannet.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(omannet.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00";
                    } else if (temp.equals("00 Dispute Server")) {
                        temp = "000 Dispute Server";
                        if (omannet.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((omannet.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(omannet.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00 Dispute Server";
                    } else {
                        if (omannet.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((omannet.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(omannet.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                    }
                   
                   
                    dataCell = dataRow.createCell(5);
                    dataCell.setCellStyle(style2);
                    if (visionPlus.get(temp) != null) {
                        dataCell.setCellValue(visionPlus.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((visionPlus.get(temp)));
                        text.append("</td>");
                    }
 
                    else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                   
                    dataCell = dataRow.createCell(6);
                    dataCell.setCellStyle(style2);
                   
                    if (temp.equals("00")) {
                        temp = "000";
                        if (amex.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((amex.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(amex.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00";
                    } else if (temp.equals("00 Dispute Server")) {
                        temp = "000 Dispute Server";
                        if (amex.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((amex.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(amex.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00 Dispute Server";
                    } else {
                        if (amex.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((amex.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(amex.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                    }
                   
                    dataCell = dataRow.createCell(7);
                    dataCell.setCellStyle(style2);
                    dataCell.setCellValue(count.get(temp));
                    text.append("<td border = '1 px solid black'>");
                    text.append(count.get(temp));
                    text.append("</td>");
                    text.append("</tr>");
 
                }
                if (tester == 0) {
                    rcText.append("<tr  border = '1 px solid black'>");
                    rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                    rcText.append("<b>-- Null --</b>" + "</td></tr>");
                }
 
                tester = 0;
 
                text.append("</table>");
                tot = sheet.getLastRowNum();
                bahrainInitial = tot + 1;
                for (int i = 0; i < tot; i++) {
                    int j = i + 1;
                    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                }
                for (int i = 1; i < 6; i++) {
                    int j = i + 1;
                    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(6, tot, i, j), sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(6, tot, i, j), sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(6, tot, i, j), sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(6, tot, i, j), sheet);
 
                }
                for (int k = 0; k < 21; k++) {
                    sheet.autoSizeColumn(k);
                }
 
                /*
                 * To maintain backup of report renaming previous report file
                 * with extention of timstamp
                 */
 
                File oldInfileName = new File(prop.getProperty("ManualomanInFile") + ".xlsx");
                File newInfileName = new File(prop.getProperty("ManualomanInFile") + timeStamp + ".xlsx");
 
                oldInfileName.renameTo(newInfileName);
                System.out.println("Rename completed successfully");
 
            }
        } catch (Exception e) {
 
            System.out.println("Exception in ABO report output...");
            fin.close();
            fo.close();
            System.out.println(e);
        }
 
        // ******************************************** bahrain code
        // *************************************************************
        try {
            if (new File(prop.getProperty("ManualbahrainInFile") + ".xlsx").exists()) {
 
                FileInputStream bahrainUpdated = new FileInputStream(prop.getProperty("ManualbahrainInFile") + ".xlsx");
                XSSFWorkbook bahrainWorkbook = new XSSFWorkbook(bahrainUpdated);
                // FileOutputStream fos = new FileOutputStream(new
                // File(prop.getProperty("bahrainBackup")));
                // bahrainWorkbook.write(fos);
                XSSFSheet bahrainSheeet = bahrainWorkbook.getSheet(bahrainWorkbook.getSheetName(0));
                TreeMap<String, Integer> bahrainCount = new TreeMap<>();
                TreeMap<String, Integer> bahrainVisa = new TreeMap<>();
                TreeMap<String, Integer> bahrainMastercard = new TreeMap<>();
                TreeMap<String, Integer> bahrainBenefit = new TreeMap<>();
                TreeMap<String, Integer> bahrainAmex = new TreeMap<>();
                TreeMap<String, Integer> bahrainVisionplus = new TreeMap<>();
                TreeMap<String, Integer> bahrainJcb = new TreeMap<>();
                TreeMap<String, Integer> bahrainOther = new TreeMap<>();
                int[] bahrainTotal = new int[7];
                // System.out.println(sheeet.getLastRowNum());
 
                rcText.append("<tr  border = '1 px solid black'>");
                rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                rcText.append("<b>========= bahrain ===========</b>" + "</td></tr>");
 
                for (int i = 8; i <= bahrainSheeet.getLastRowNum() - 2; i++) {
                    XSSFRow row = bahrainSheeet.getRow(i);
                    XSSFCell cell = row.getCell(18);
                    String code = cell.getStringCellValue();
                    XSSFCell cell2 = row.getCell(17);
                    String interchange = cell2.getStringCellValue();
                    cell2 = row.getCell(19);
                    String status = cell2.getStringCellValue();
                    if ((code.equals("") || code.equals(null) || code.equals(" ") || code.equals("0"))) {
                        if (row.getCell(20).getStringCellValue().equalsIgnoreCase("In progress")) {
                            code = new String("In progress");
                        } else {
                            code = row.getCell(20).getStringCellValue();
                        }
                    } else if (row.getCell(20).getStringCellValue().equalsIgnoreCase("Timeout")) {
 
                        code = code + " " + new String(row.getCell(20).getStringCellValue());
 
                    } else if (!(status.equals("Not initiated") || status.equals("Settled"))) {
                        System.out.println("here" + status);
                        System.out.println("RRn    " + row.getCell(6).getStringCellValue());
                        code = code + " " + status;
                        System.out.println("code");
                    }
 
                    switch (interchange) {
                    case "BENEFIT":
                        bahrainTotal[2]++;
                        if (bahrainBenefit.containsKey(code))
                            bahrainBenefit.put(code, bahrainBenefit.get(code) + 1);
                        else
                            bahrainBenefit.put(code, 1);
                        break;
                    case "MASTERCARD":
                        bahrainTotal[1]++;
                        if (bahrainMastercard.containsKey(code))
                            bahrainMastercard.put(code, bahrainMastercard.get(code) + 1);
                        else
                            bahrainMastercard.put(code, 1);
                        break;
                    case "VISA":
                        bahrainTotal[0]++;
                        if (bahrainVisa.containsKey(code))
                            bahrainVisa.put(code, bahrainVisa.get(code) + 1);
                        else
                            bahrainVisa.put(code, 1);
                        break;
                    case "AMEXACQUIRER":
                        bahrainTotal[3]++;
                        if (bahrainAmex.containsKey(code))
                            bahrainAmex.put(code, bahrainAmex.get(code) + 1);
                        else
                            bahrainAmex.put(code, 1);
                        break;
                    case "VISIONPLUSHOST":
                        bahrainTotal[4]++;
                        if (bahrainVisionplus.containsKey(code))
                            bahrainVisionplus.put(code, bahrainVisionplus.get(code) + 1);
                        else
                            bahrainVisionplus.put(code, 1);
                        break;
                    case "JCB":
                        bahrainTotal[5]++;
                        if (bahrainJcb.containsKey(code))
                            bahrainJcb.put(code, bahrainJcb.get(code) + 1);
                        else
                            bahrainJcb.put(code, 1);
                        break;
                    default:
                        if (bahrainOther.containsKey(code))
                            bahrainOther.put(code, bahrainOther.get(code) + 1);
                        else
                            bahrainOther.put(code, 1);
                    }
 
                    if (bahrainCount.containsKey(code)) {
                        bahrainCount.put(code, bahrainCount.get(code) + 1);
                    } else {
                        bahrainCount.put(code, 1);
                    }
                    bahrainTotal[6]++;
                }
                System.out.println("bahrain Count: " + bahrainCount);
 
                int bahrainSuccess = 0;
 
                if (bahrainCount.get("000") != null && bahrainCount.get("00") != null) {
                    bahrainSuccess = bahrainCount.get("00") + bahrainCount.get("000");
                    bahrainCount.put("00", bahrainSuccess);
                    bahrainCount.remove("000");
                } else if (bahrainCount.get("000") == null && bahrainCount.get("00") == null) {
 
                } else {
                    if (bahrainCount.get("00") != null) {
                        bahrainSuccess = bahrainCount.get("00");
                    } else
                        bahrainSuccess = bahrainCount.get("000");
                    bahrainCount.put("00", bahrainSuccess);
                    bahrainCount.remove("000");
                }
 
                System.out.println(bahrainCount);
                HSSFRow bahrainHeaderRow = sheet.createRow(sheet.getLastRowNum() + 2);
                HSSFCell bahrainHeaderCell = bahrainHeaderRow.createCell(0);
                bahrainHeaderCell.setCellValue("Transactions Detail report - bahrain");
                sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 8));
                bahrainHeaderCell.setCellStyle(style);
                text.append("<br>");
                text.append(
                        "<table border='1'  border = '1 px solid black' bordercolor='BLACK' style='border-collapse:collapse; font-family:Calibri;'>");
                text.append("<tr  border = '1 px solid black'>");
                text.append("<td colspan='9'  border = '1 px solid black'style ='text-align : center'>");
                text.append("<b>Transactions Detail report - bahrain</b>");
                text.append("</td>");
                text.append("</tr>");
                sheet.autoSizeColumn(0);
                bahrainHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
 
                // Makarand code for new output Sheet
                String bankbahrain = "Financial Institution ID : ALI - Ahli United Bank";
                String tranDatebahrain = "Transaction Date : " + getCurrentBahrainTime(); // append
                                                                                        // Bahrain
                                                                                        // Time
                String runDatebahrain = "Run Date/Time : " + java.time.LocalDate.now() + " " + getCurrentBahrainTime()
                        + " Asia/Muscat"; // Append IST Date and Oman Time
 
                for (int i = 1; i < 4; i++) {
                    bahrainHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    HSSFCell bankCell = bahrainHeaderRow.createCell(0);
                    sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 0, 8));
                    bankCell.setCellStyle(style1);
 
                    if (i == 1) {
 
                        bankCell.setCellValue(bankbahrain);
                        // sheet.addMergedRegion(new
                        // CellRangeAddress(sheet.getLastRowNum(),
                        // sheet.getLastRowNum(), 0, 8));
                        // bankCell.setCellStyle(style1);
                        text.append("<tr>");
                        text.append("<td colspan='9'   border = '1 px solid black'>");
                        text.append("<b>" + bankbahrain + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                    if (i == 2) {
 
                        bankCell.setCellValue(tranDatebahrain);
                        // sheet.addMergedRegion(new
                        // CellRangeAddress(sheet.getLastRowNum(),
                        // sheet.getLastRowNum(), 0, 8));
                        // bankCell.setCellStyle(style1);
                        text.append("<tr>");
                        text.append("<td colspan='9'   border = '1 px solid black'>");
                        text.append("<b>" + tranDatebahrain + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                    if (i == 3) {
 
                        bankCell.setCellValue(runDatebahrain);
                        // sheet.addMergedRegion(new
                        // CellRangeAddress(sheet.getLastRowNum(),
                        // sheet.getLastRowNum(), 0, 8));
                        // bankCell.setCellStyle(style1);
                        text.append("<tr>");
                        text.append("<td colspan='9'   border = '1 px solid black'>");
                        text.append("<b>" + runDatebahrain + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                }
 
                /*
                 * for (int i = 2; i <= 5; i++) { bahrainHeaderRow =
                 * sheet.createRow(sheet.getLastRowNum() + 1); HSSFCell bankCell
                 * = bahrainHeaderRow.createCell(0); String data; if (i == 5) {
                 * String tempData =
                 * bahrainSheeet.getRow(i).getCell(0).getStringCellValue(); //
                 * System.out.println(tempData.lastIndexOf('/'));
                 *
                 * data = tempData.substring(0, tempData.lastIndexOf('/')); //
                 * System.out.println(data.concat("/Muscat")); data = data +
                 * "/Muscat";
                 *
                 * } else { data =
                 * bahrainSheeet.getRow(i).getCell(0).getStringCellValue(); } //
                 * System.out.println(data + " " + i);
                 * bankCell.setCellValue(data); sheet.addMergedRegion(new
                 * CellRangeAddress(sheet.getLastRowNum(),
                 * sheet.getLastRowNum(), 0, 8)); bankCell.setCellStyle(style1);
                 * bankCell.setCellValue(data); text.append("<tr>");
                 * text.append("<td colspan='9'   border = '1 px solid black'>"
                 * ); text.append("<b>" + data + "</b>"); text.append("</td>");
                 * text.append("</tr>"); }
                 */
 
                HSSFRow bahrainRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                HSSFCell bahrainCellTotal = bahrainRowTotal.createCell(0);
                bahrainCellTotal.setCellStyle(style);
                sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 1));
                bahrainCellTotal.setCellValue("Total Transactions");
                text.append("<tr>");
                text.append("<td border = '1 px solid black' colspan='2' rowspan='2'style ='text-align : center'>");
                text.append("<b>Total Transactions</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(2);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("VISA");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISA</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(3);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("MASTER CARD");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Master Card</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(4);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("BENEFIT");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>BENEFIT</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(5);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("AMEX");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>AMEX</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(6);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("VISION PLUS");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISION PLUS</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(7);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("JCB");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>JCB</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(8);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("GRAND TOTAL");
                bahrainRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                text.append("<td border = '1 px solid black'>");
                text.append("<b>GRAND TOTAL</b>");
                text.append("</td>");
                text.append("</tr>");
                text.append("<tr>");
                for (int i = 0; i < bahrainTotal.length; i++) {
                    int j = i + 2;
                    bahrainCellTotal = bahrainRowTotal.createCell(j);
                    bahrainCellTotal.setCellStyle(style2);
                    bahrainCellTotal.setCellValue(bahrainTotal[i]);
                    text.append("<td  border = '1 px solid black'>");
                    text.append(bahrainTotal[i]);
                    text.append("</td>");
                }
 
                text.append("</tr>");
                text.append("<tr>");
                HSSFRow bahrainRowBlank = sheet.createRow(sheet.getLastRowNum() + 1);
                HSSFCell bahrainCellHeader = bahrainRowBlank.createCell(0);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("Response code");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Response code</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(1);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("Response code description");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Response code description/<b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(2);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("VISA");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISA</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(3);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("MASTERCARD");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>MASTERCARD</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(4);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("BENEFIT");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>BENEFIT</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(5);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("AMEX");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>AMEX</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(6);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("VISION PLUS");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISION PLUS</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(7);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("JCB");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>JCB</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(8);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("Grand Total");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Grand Total</b>");
                text.append("</td>");
                text.append("</tr>");
 
                sheet.autoSizeColumn(0);
                sheet.autoSizeColumn(1);
                sheet.autoSizeColumn(2);
                sheet.autoSizeColumn(3);
                sheet.autoSizeColumn(4);
                sheet.autoSizeColumn(5);
                sheet.autoSizeColumn(6);
                sheet.autoSizeColumn(7);
                sheet.autoSizeColumn(8);
                int bahrainLength = sheet.getLastRowNum() + 1;
 
                Set<String> bahrainKeys = bahrainCount.keySet();
                Iterator<String> bahrainIt = bahrainKeys.iterator();
                String settlementStatus = "";
                for (int i = 2; i <= bahrainCount.size() + 1; i++) {
                    text.append("<tr>");
                    String temp = bahrainIt.next();
                    HSSFRow dataRow = sheet.createRow(bahrainLength++);
                    dataRow.setRowStyle(style);
                    HSSFCell dataCell = dataRow.createCell(0);
                    dataCell.setCellStyle(style2);
                    dataCell.setCellValue(temp);
                    text.append("<td border = '1 px solid black'>");
                    text.append(temp);
                    text.append("</td>");
 
                    if (temp.length() > 4) {
                        settlementStatus = temp.substring(3, temp.length());
                        System.out.println(settlementStatus);
                    }
 
                    dataCell = dataRow.createCell(1);
                    dataCell.setCellStyle(style2);
 
                    if (prop.getProperty(temp) != null) {
                        dataCell.setCellValue(prop.getProperty(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append(prop.getProperty(temp));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue("unknown response code/ unknown status");
                        text.append("<td border = '1 px solid black'>");
                        text.append("unknown response code/ unknown status");
                        text.append("</td>");
                    }
                    // }
                    dataCell = dataRow.createCell(2);
                    dataCell.setCellStyle(style2);
                    if (bahrainVisa.get(temp) != null) {
                        dataCell.setCellValue(bahrainVisa.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((bahrainVisa.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(3);
                    dataCell.setCellStyle(style2);
                    if (bahrainMastercard.get(temp) != null) {
                        dataCell.setCellValue(bahrainMastercard.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((bahrainMastercard.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(4);
                    dataCell.setCellStyle(style2);
                    if (bahrainBenefit.get(temp) != null) {
                        dataCell.setCellValue(bahrainBenefit.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((bahrainBenefit.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(5);
                    dataCell.setCellStyle(style2);
                    if (temp.equals("00")) {
                        temp = "000";
                        if (bahrainAmex.get(temp) != null) {
                            dataCell.setCellValue(bahrainAmex.get(temp));
                            text.append("<td border = '1 px solid black'>");
                            text.append((bahrainAmex.get(temp)));
                            text.append("</td>");
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00";
                    } else if (temp.equals("00 Dispute Server")) {
                        if (bahrainAmex.get(temp) != null) {
                            dataCell.setCellValue(bahrainAmex.get(temp));
                            text.append("<td border = '1 px solid black'>");
                            text.append((bahrainAmex.get(temp)));
                            text.append("</td>");
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00 Dispute Server";
                    } else {
                        if (bahrainAmex.get(temp) != null) {
                            dataCell.setCellValue(bahrainAmex.get(temp));
                            text.append("<td border = '1 px solid black'>");
                            text.append((bahrainAmex.get(temp)));
                            text.append("</td>");
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                    }
                    dataCell = dataRow.createCell(6);
                    dataCell.setCellStyle(style2);
                    if (bahrainVisionplus.get(temp) != null) {
                        dataCell.setCellValue(bahrainVisionplus.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((bahrainVisionplus.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(7);
                    dataCell.setCellStyle(style2);
                    if (bahrainJcb.get(temp) != null) {
                        dataCell.setCellValue(bahrainJcb.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((bahrainJcb.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(8);
                    dataCell.setCellStyle(style2);
                    dataCell.setCellValue(bahrainCount.get(temp));
                    text.append("<td border = '1 px solid black'>");
                    text.append(bahrainCount.get(temp));
                    text.append("</td>");
                    text.append("</tr>");
 
                }
 
                /*
                 * if(tester == 0) { rcText.append(
                 * "<tr  border = '1 px solid black'>"); rcText.append(
                 * "<td colspan='3'  border = '1 px solid black'style ='text-align : center'>"
                 * ); rcText.append("<b>-- Null --</b>"+ "</td></tr>"); }
                 */
                tester = 0;
 
                text.append("</table>");
                text.append("<br>");
                tot = sheet.getLastRowNum();
                for (int i = bahrainInitial + 2; i < tot; i++) {
                    int j = i + 1;
                    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                }
                for (int i = 0; i < 8; i++) {
                    int j = i + 1;
                    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(bahrainInitial + 7, tot, i, j),
                            sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(bahrainInitial + 7, tot, i, j), sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(bahrainInitial + 7, tot, i, j), sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(bahrainInitial + 7, tot, i, j), sheet);
 
                }
 
                /*
                 * To maintain backup of report renaming previous report file
                 * with extention of timstamp
                 */
 
                File oldInfileName = new File(prop.getProperty("ManualbahrainInFile") + ".xlsx");
                File newInfileName = new File(prop.getProperty("ManualbahrainInFile") + timeStamp + ".xlsx");
 
                oldInfileName.renameTo(newInfileName);
                System.out.println("Rename completed successfully");
 
                // *************** Code for comparing grandtotal for validating
                // GG issue. ***********************
 
                // prop.setProperty("bahrainTotalCurr",Integer.toString(bahrainTotal[6]));
 
                // ************** End of Code for comparing grandtotal for
                // validating GG issue *******************
 
            } else {
                System.out.println("bahrain Infile not present");
            }
        } catch (Exception e) {
 
            System.out.println("Exception in bahrain report output...");
            fin.close();
            fo.close();
            System.out.println(e);
        }
 
        // ********************************************* AUB CODE
        // *************************************************************
        try {
            if (new File(prop.getProperty("ManualaubInFile") + ".xlsx").exists()) {
 
                FileInputStream AUBUpdated = new FileInputStream(prop.getProperty("ManualaubInFile") + ".xlsx");
                XSSFWorkbook AUBWorkbook = new XSSFWorkbook(AUBUpdated);
                // FileOutputStream fos = new FileOutputStream(new
                // File(prop.getProperty("aubBackup")));
                // AUBWorkbook.write(fos);
                XSSFSheet AUBSheeet = AUBWorkbook.getSheet(AUBWorkbook.getSheetName(0));
                TreeMap<String, Integer> AUBCount = new TreeMap<>();
                TreeMap<String, Integer> AUBVisa = new TreeMap<>();
                TreeMap<String, Integer> AUBMastercard = new TreeMap<>();
                TreeMap<String, Integer> AUBBenefit = new TreeMap<>();
                TreeMap<String, Integer> AUBAmex = new TreeMap<>();
                TreeMap<String, Integer> AUBVisionplus = new TreeMap<>();
                TreeMap<String, Integer> AUBJcb = new TreeMap<>();
                TreeMap<String, Integer> AUBOther = new TreeMap<>();
                int[] AUBTotal = new int[7];
                // System.out.println(sheeet.getLastRowNum());
 
                rcText.append("<tr  border = '1 px solid black'>");
                rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                rcText.append("<b>========= AUB ===========</b>" + "</td></tr>");
 
                for (int i = 8; i <= AUBSheeet.getLastRowNum() - 2; i++) {
                    XSSFRow row = AUBSheeet.getRow(i);
                    XSSFCell cell = row.getCell(18);
                    String code = cell.getStringCellValue();
                    XSSFCell cell2 = row.getCell(17);
                    String interchange = cell2.getStringCellValue();
                    cell2 = row.getCell(19);
                    String status = cell2.getStringCellValue();
                    if ((code.equals("") || code.equals(null) || code.equals(" ") || code.equals("0"))) {
                        if (row.getCell(20).getStringCellValue().equalsIgnoreCase("In progress")) {
                            code = new String("In progress");
                        } else {
                            code = row.getCell(20).getStringCellValue();
                        }
                    } else if (row.getCell(20).getStringCellValue().equalsIgnoreCase("Timeout")) {
 
                        code = code + " " + new String(row.getCell(20).getStringCellValue());
 
                    } else if (!(status.equals("Not initiated") || status.equals("Settled"))) {
                        System.out.println("here" + status);
                        System.out.println("RRn    " + row.getCell(6).getStringCellValue());
                        code = code + " " + status;
                        System.out.println("code");
                    }
 
                    switch (interchange) {
                    case "BENEFIT":
                        AUBTotal[2]++;
                        if (AUBBenefit.containsKey(code))
                            AUBBenefit.put(code, AUBBenefit.get(code) + 1);
                        else
                            AUBBenefit.put(code, 1);
                        break;
                    case "MASTER CARD":
                        AUBTotal[1]++;
                        if (AUBMastercard.containsKey(code))
                            AUBMastercard.put(code, AUBMastercard.get(code) + 1);
                        else
                            AUBMastercard.put(code, 1);
                        break;
                    case "VISA":
                        AUBTotal[0]++;
                        if (AUBVisa.containsKey(code))
                            AUBVisa.put(code, AUBVisa.get(code) + 1);
                        else
                            AUBVisa.put(code, 1);
                        break;
                    case "AMEX":
                        AUBTotal[3]++;
                        if (AUBAmex.containsKey(code))
                            AUBAmex.put(code, AUBAmex.get(code) + 1);
                        else
                            AUBAmex.put(code, 1);
                        break;
                    case "VISIONPLUSHOST":
                        AUBTotal[4]++;
                        if (AUBVisionplus.containsKey(code))
                            AUBVisionplus.put(code, AUBVisionplus.get(code) + 1);
                        else
                            AUBVisionplus.put(code, 1);
                        break;
                    case "JCB":
                        AUBTotal[5]++;
                        if (AUBJcb.containsKey(code))
                            AUBJcb.put(code, AUBJcb.get(code) + 1);
                        else
                            AUBJcb.put(code, 1);
                        break;
                    default:
                        if (AUBOther.containsKey(code))
                            AUBOther.put(code, AUBOther.get(code) + 1);
                        else
                            AUBOther.put(code, 1);
                    }
 
                    if (AUBCount.containsKey(code)) {
                        AUBCount.put(code, AUBCount.get(code) + 1);
                    } else {
                        AUBCount.put(code, 1);
                    }
                    AUBTotal[6]++;
                }
                System.out.println("AUB Count: " + AUBCount);
 
                int AUBSuccess = 0;
 
                if (AUBCount.get("000") != null && AUBCount.get("00") != null) {
                    AUBSuccess = AUBCount.get("00") + AUBCount.get("000");
                    AUBCount.put("00", AUBSuccess);
                    AUBCount.remove("000");
                } else if (AUBCount.get("000") == null && AUBCount.get("00") == null) {
 
                } else {
                    if (AUBCount.get("00") != null) {
                        AUBSuccess = AUBCount.get("00");
                    } else
                        AUBSuccess = AUBCount.get("000");
                    AUBCount.put("00", AUBSuccess);
                    AUBCount.remove("000");
                }
 
                System.out.println(AUBCount);
                HSSFRow AUBHeaderRow = sheet.createRow(sheet.getLastRowNum() + 2);
                HSSFCell AUBHeaderCell = AUBHeaderRow.createCell(0);
                AUBHeaderCell.setCellValue("Transactions Detail report - AUB");
                sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 8));
                AUBHeaderCell.setCellStyle(style);
                text.append("<br>");
                text.append(
                        "<table border='1'  border = '1 px solid black' bordercolor='BLACK' style='border-collapse:collapse; font-family:Calibri;'>");
                text.append("<tr  border = '1 px solid black'>");
                text.append("<td colspan='9'  border = '1 px solid black'style ='text-align : center'>");
                text.append("<b>Transactions Detail report - AUB</b>");
                text.append("</td>");
                text.append("</tr>");
                sheet.autoSizeColumn(0);
                AUBHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
 
                // Makarand code for new output Sheet
                String bankAUB = "Financial Institution ID : ALI - Ahli United Bank";
                String tranDateAUB = "Transaction Date : " + getCurrentBahrainTime(); // append
                                                                                        // Bahrain
                                                                                        // Time
                String runDateAUB = "Run Date/Time : " + java.time.LocalDate.now() + " " + getCurrentBahrainTime()
                        + " Asia/Muscat"; // Append IST Date and Oman Time
 
                for (int i = 1; i < 4; i++) {
                    AUBHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    HSSFCell bankCell = AUBHeaderRow.createCell(0);
                    sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 0, 8));
                    bankCell.setCellStyle(style1);
 
                    if (i == 1) {
 
                        bankCell.setCellValue(bankAUB);
                        // sheet.addMergedRegion(new
                        // CellRangeAddress(sheet.getLastRowNum(),
                        // sheet.getLastRowNum(), 0, 8));
                        // bankCell.setCellStyle(style1);
                        text.append("<tr>");
                        text.append("<td colspan='9'   border = '1 px solid black'>");
                        text.append("<b>" + bankAUB + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                    if (i == 2) {
 
                        bankCell.setCellValue(tranDateAUB);
                        // sheet.addMergedRegion(new
                        // CellRangeAddress(sheet.getLastRowNum(),
                        // sheet.getLastRowNum(), 0, 8));
                        // bankCell.setCellStyle(style1);
                        text.append("<tr>");
                        text.append("<td colspan='9'   border = '1 px solid black'>");
                        text.append("<b>" + tranDateAUB + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                    if (i == 3) {
 
                        bankCell.setCellValue(runDateAUB);
                        // sheet.addMergedRegion(new
                        // CellRangeAddress(sheet.getLastRowNum(),
                        // sheet.getLastRowNum(), 0, 8));
                        // bankCell.setCellStyle(style1);
                        text.append("<tr>");
                        text.append("<td colspan='9'   border = '1 px solid black'>");
                        text.append("<b>" + runDateAUB + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                }
 
                /*
                 * for (int i = 2; i <= 5; i++) { AUBHeaderRow =
                 * sheet.createRow(sheet.getLastRowNum() + 1); HSSFCell bankCell
                 * = AUBHeaderRow.createCell(0); String data; if (i == 5) {
                 * String tempData =
                 * AUBSheeet.getRow(i).getCell(0).getStringCellValue(); //
                 * System.out.println(tempData.lastIndexOf('/'));
                 *
                 * data = tempData.substring(0, tempData.lastIndexOf('/')); //
                 * System.out.println(data.concat("/Muscat")); data = data +
                 * "/Muscat";
                 *
                 * } else { data =
                 * AUBSheeet.getRow(i).getCell(0).getStringCellValue(); } //
                 * System.out.println(data + " " + i);
                 * bankCell.setCellValue(data); sheet.addMergedRegion(new
                 * CellRangeAddress(sheet.getLastRowNum(),
                 * sheet.getLastRowNum(), 0, 8)); bankCell.setCellStyle(style1);
                 * bankCell.setCellValue(data); text.append("<tr>");
                 * text.append("<td colspan='9'   border = '1 px solid black'>"
                 * ); text.append("<b>" + data + "</b>"); text.append("</td>");
                 * text.append("</tr>"); }
                 */
 
                HSSFRow AUBRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                HSSFCell AUBCellTotal = AUBRowTotal.createCell(0);
                AUBCellTotal.setCellStyle(style);
                sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 1));
                AUBCellTotal.setCellValue("Total Transactions");
                text.append("<tr>");
                text.append("<td border = '1 px solid black' colspan='2' rowspan='2'style ='text-align : center'>");
                text.append("<b>Total Transactions</b>");
                text.append("</td>");
                AUBCellTotal = AUBRowTotal.createCell(2);
                AUBCellTotal.setCellStyle(style);
                AUBCellTotal.setCellValue("VISA");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISA</b>");
                text.append("</td>");
                AUBCellTotal = AUBRowTotal.createCell(3);
                AUBCellTotal.setCellStyle(style);
                AUBCellTotal.setCellValue("MASTER CARD");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Master Card</b>");
                text.append("</td>");
                AUBCellTotal = AUBRowTotal.createCell(4);
                AUBCellTotal.setCellStyle(style);
                AUBCellTotal.setCellValue("BENEFIT");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>BENEFIT</b>");
                text.append("</td>");
                AUBCellTotal = AUBRowTotal.createCell(5);
                AUBCellTotal.setCellStyle(style);
                AUBCellTotal.setCellValue("AMEX");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>AMEX</b>");
                text.append("</td>");
                AUBCellTotal = AUBRowTotal.createCell(6);
                AUBCellTotal.setCellStyle(style);
                AUBCellTotal.setCellValue("VISION PLUS");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISION PLUS</b>");
                text.append("</td>");
                AUBCellTotal = AUBRowTotal.createCell(7);
                AUBCellTotal.setCellStyle(style);
                AUBCellTotal.setCellValue("JCB");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>JCB</b>");
                text.append("</td>");
                AUBCellTotal = AUBRowTotal.createCell(8);
                AUBCellTotal.setCellStyle(style);
                AUBCellTotal.setCellValue("GRAND TOTAL");
                AUBRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                text.append("<td border = '1 px solid black'>");
                text.append("<b>GRAND TOTAL</b>");
                text.append("</td>");
                text.append("</tr>");
                text.append("<tr>");
                for (int i = 0; i < AUBTotal.length; i++) {
                    int j = i + 2;
                    AUBCellTotal = AUBRowTotal.createCell(j);
                    AUBCellTotal.setCellStyle(style2);
                    AUBCellTotal.setCellValue(AUBTotal[i]);
                    text.append("<td  border = '1 px solid black'>");
                    text.append(AUBTotal[i]);
                    text.append("</td>");
                }
 
                text.append("</tr>");
                text.append("<tr>");
                HSSFRow AUBRowBlank = sheet.createRow(sheet.getLastRowNum() + 1);
                HSSFCell AUBCellHeader = AUBRowBlank.createCell(0);
                AUBCellHeader.setCellStyle(style);
                AUBCellHeader.setCellValue("Response code");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Response code</b>");
                text.append("</td>");
                AUBCellHeader = AUBRowBlank.createCell(1);
                AUBCellHeader.setCellStyle(style);
                AUBCellHeader.setCellValue("Response code description");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Response code description/<b>");
                text.append("</td>");
                AUBCellHeader = AUBRowBlank.createCell(2);
                AUBCellHeader.setCellStyle(style);
                AUBCellHeader.setCellValue("VISA");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISA</b>");
                text.append("</td>");
                AUBCellHeader = AUBRowBlank.createCell(3);
                AUBCellHeader.setCellStyle(style);
                AUBCellHeader.setCellValue("MASTERCARD");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>MASTERCARD</b>");
                text.append("</td>");
                AUBCellHeader = AUBRowBlank.createCell(4);
                AUBCellHeader.setCellStyle(style);
                AUBCellHeader.setCellValue("BENEFIT");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>BENEFIT</b>");
                text.append("</td>");
                AUBCellHeader = AUBRowBlank.createCell(5);
                AUBCellHeader.setCellStyle(style);
                AUBCellHeader.setCellValue("AMEX");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>AMEX</b>");
                text.append("</td>");
                AUBCellHeader = AUBRowBlank.createCell(6);
                AUBCellHeader.setCellStyle(style);
                AUBCellHeader.setCellValue("VISION PLUS");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISION PLUS</b>");
                text.append("</td>");
                AUBCellHeader = AUBRowBlank.createCell(7);
                AUBCellHeader.setCellStyle(style);
                AUBCellHeader.setCellValue("JCB");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>JCB</b>");
                text.append("</td>");
                AUBCellHeader = AUBRowBlank.createCell(8);
                AUBCellHeader.setCellStyle(style);
                AUBCellHeader.setCellValue("Grand Total");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Grand Total</b>");
                text.append("</td>");
                text.append("</tr>");
 
                sheet.autoSizeColumn(0);
                sheet.autoSizeColumn(1);
                sheet.autoSizeColumn(2);
                sheet.autoSizeColumn(3);
                sheet.autoSizeColumn(4);
                sheet.autoSizeColumn(5);
                sheet.autoSizeColumn(6);
                sheet.autoSizeColumn(7);
                sheet.autoSizeColumn(8);
                int AUBLength = sheet.getLastRowNum() + 1;
 
                Set<String> AUBKeys = AUBCount.keySet();
                Iterator<String> AUBIt = AUBKeys.iterator();
                String settlementStatus = "";
                for (int i = 2; i <= AUBCount.size() + 1; i++) {
                    text.append("<tr>");
                    String temp = AUBIt.next();
                    HSSFRow dataRow = sheet.createRow(AUBLength++);
                    dataRow.setRowStyle(style);
                    HSSFCell dataCell = dataRow.createCell(0);
                    dataCell.setCellStyle(style2);
                    dataCell.setCellValue(temp);
                    text.append("<td border = '1 px solid black'>");
                    text.append(temp);
                    text.append("</td>");
 
                    if (temp.length() > 4) {
                        settlementStatus = temp.substring(3, temp.length());
                        System.out.println(settlementStatus);
                    }
 
                    dataCell = dataRow.createCell(1);
                    dataCell.setCellStyle(style2);
 
                    if (prop.getProperty(temp) != null) {
                        dataCell.setCellValue(prop.getProperty(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append(prop.getProperty(temp));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue("unknown response code/ unknown status");
                        text.append("<td border = '1 px solid black'>");
                        text.append("unknown response code/ unknown status");
                        text.append("</td>");
                    }
                    // }
                    dataCell = dataRow.createCell(2);
                    dataCell.setCellStyle(style2);
                    if (AUBVisa.get(temp) != null) {
                        dataCell.setCellValue(AUBVisa.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((AUBVisa.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(3);
                    dataCell.setCellStyle(style2);
                    if (AUBMastercard.get(temp) != null) {
                        dataCell.setCellValue(AUBMastercard.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((AUBMastercard.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(4);
                    dataCell.setCellStyle(style2);
                    if (AUBBenefit.get(temp) != null) {
                        dataCell.setCellValue(AUBBenefit.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((AUBBenefit.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(5);
                    dataCell.setCellStyle(style2);
                    if (temp.equals("00")) {
                        temp = "000";
                        if (AUBAmex.get(temp) != null) {
                            dataCell.setCellValue(AUBAmex.get(temp));
                            text.append("<td border = '1 px solid black'>");
                            text.append((AUBAmex.get(temp)));
                            text.append("</td>");
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00";
                    } else if (temp.equals("00 Dispute Server")) {
                        if (AUBAmex.get(temp) != null) {
                            dataCell.setCellValue(AUBAmex.get(temp));
                            text.append("<td border = '1 px solid black'>");
                            text.append((AUBAmex.get(temp)));
                            text.append("</td>");
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00 Dispute Server";
                    } else {
                        if (AUBAmex.get(temp) != null) {
                            dataCell.setCellValue(AUBAmex.get(temp));
                            text.append("<td border = '1 px solid black'>");
                            text.append((AUBAmex.get(temp)));
                            text.append("</td>");
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                    }
                    dataCell = dataRow.createCell(6);
                    dataCell.setCellStyle(style2);
                    if (AUBVisionplus.get(temp) != null) {
                        dataCell.setCellValue(AUBVisionplus.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((AUBVisionplus.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(7);
                    dataCell.setCellStyle(style2);
                    if (AUBJcb.get(temp) != null) {
                        dataCell.setCellValue(AUBJcb.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((AUBJcb.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(8);
                    dataCell.setCellStyle(style2);
                    dataCell.setCellValue(AUBCount.get(temp));
                    text.append("<td border = '1 px solid black'>");
                    text.append(AUBCount.get(temp));
                    text.append("</td>");
                    text.append("</tr>");
 
                }
 
                /*
                 * if(tester == 0) { rcText.append(
                 * "<tr  border = '1 px solid black'>"); rcText.append(
                 * "<td colspan='3'  border = '1 px solid black'style ='text-align : center'>"
                 * ); rcText.append("<b>-- Null --</b>"+ "</td></tr>"); }
                 */
                tester = 0;
 
                text.append("</table>");
                text.append("<br>");
                tot = sheet.getLastRowNum();
                for (int i = AUBInitial + 2; i < tot; i++) {
                    int j = i + 1;
                    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                }
                for (int i = 0; i < 8; i++) {
                    int j = i + 1;
                    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(AUBInitial + 7, tot, i, j),
                            sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(AUBInitial + 7, tot, i, j), sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(AUBInitial + 7, tot, i, j), sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(AUBInitial + 7, tot, i, j), sheet);
 
                }
 
                /*
                 * To maintain backup of report renaming previous report file
                 * with extention of timstamp
                 */
 
                File oldInfileName = new File(prop.getProperty("ManualaubInFile") + ".xlsx");
                File newInfileName = new File(prop.getProperty("ManualaubInFile") + timeStamp + ".xlsx");
 
                oldInfileName.renameTo(newInfileName);
                System.out.println("Rename completed successfully");
 
                // *************** Code for comparing grandtotal for validating
                // GG issue. ***********************
 
                // prop.setProperty("aubTotalCurr",Integer.toString(AUBTotal[6]));
 
                // ************** End of Code for comparing grandtotal for
                // validating GG issue *******************
 
            } else {
                System.out.println("AUB Infile not present");
            }
        } catch (Exception e) {
 
            System.out.println("Exception in AUB report output...");
            fin.close();
            fo.close();
            System.out.println(e);
        }
 
        // ****************************** Sohar Environment***********************************
       
        try {
            if (new File(prop.getProperty("ManualsoharInFile") + ".xlsx").exists()) {
                TreeMap<String, Integer> soharCount = new TreeMap<>();
                TreeMap<String, Integer> soharVisa = new TreeMap<>();
                TreeMap<String, Integer> soharMastercard = new TreeMap<>();
                TreeMap<String, Integer> soharHpsSohar = new TreeMap<>();
                TreeMap<String, Integer> soharVisionPlus = new TreeMap<>();
                TreeMap<String, Integer> soharAmex = new TreeMap<>();
                TreeMap<String, Integer> soharOther = new TreeMap<>();
                int[] soharTotal = new int[6];
 
                FileInputStream soharUpdated = new FileInputStream(prop.getProperty("ManualsoharInFile") + ".xlsx");
                XSSFWorkbook soharWorkbook = new XSSFWorkbook(soharUpdated);
                // FileOutputStream fos = new FileOutputStream(new
                // File(prop.getProperty("soharBackup")));
                // soharWorkbook.write(fos);
                XSSFSheet soharSheeet = soharWorkbook.getSheet(soharWorkbook.getSheetName(0));
 
                rcText.append("<tr  border = '1 px solid black'>");
                rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                rcText.append("<b>========= BSH ===========</b>" + "</td></tr>");
 
                for (int i = 8; i <= soharSheeet.getLastRowNum() - 2; i++) { // initial
                                                                                // i=8
                    XSSFRow soharRow = soharSheeet.getRow(i);
                    /*
                     * System.out.println(
                     * "**********************************************");
                     * System.out.print("BSH cell 17"
                     * +soharRow.getCell(17).getStringCellValue());
                     * System.out.print("BSH cell 18"
                     * +soharRow.getCell(18).getStringCellValue());
                     * System.out.print("BSH cell 19"
                     * +soharRow.getCell(19).getStringCellValue());
                     * System.out.print("BSH cell 20"
                     * +soharRow.getCell(20).getStringCellValue());
                     * System.out.println(
                     * "**********************************************");
                     */
 
                    XSSFCell cell = soharRow.getCell(18); // Resp Code
                    String code = cell.getStringCellValue();
                    XSSFCell cell2 = soharRow.getCell(17); // Interchange
                    String interchange = cell2.getStringCellValue();
                    cell2 = soharRow.getCell(19); // Settlement Status ; col 19
                                                    // = Tran status
                    String status = cell2.getStringCellValue();
 
                    if ((code.equals("") || code.equals(null) || code.equals(" ") || code.equals("0"))) {
                        if (soharRow.getCell(20).getStringCellValue().equalsIgnoreCase("In progress")) {
                            code = new String("In progress");
                        } else {
                            code = soharRow.getCell(20).getStringCellValue();
                        }
                    } else if (soharRow.getCell(20).getStringCellValue().equalsIgnoreCase("Timeout")) {
 
                        code = code + " " + new String(soharRow.getCell(20).getStringCellValue());
 
                    } else if (!(status.equals("Not initiated") || status.equals("Settled"))) {
                        code = code + " " + status;
                    }
 
                    switch (interchange) {
                    case "HPS SOHAR":
                        soharTotal[2]++;
                        if (soharHpsSohar.containsKey(code))
                            soharHpsSohar.put(code, soharHpsSohar.get(code) + 1);
                        else
                            soharHpsSohar.put(code, 1);
                        break;
                    case "MASTERCARD":
                        soharTotal[1]++;
                        if (soharMastercard.containsKey(code))
                            soharMastercard.put(code, soharMastercard.get(code) + 1);
                        else
                            soharMastercard.put(code, 1);
                        break;
                    case "VISA":
                        soharTotal[0]++;
                        if (soharVisa.containsKey(code))
                            soharVisa.put(code, soharVisa.get(code) + 1);
                        else
                            soharVisa.put(code, 1);
                        break;
                    case "VISION PLUS":
                        soharTotal[3]++;
                        if (soharVisionPlus.containsKey(code))
                            soharVisionPlus.put(code, soharVisionPlus.get(code) + 1);
                        else
                            soharVisionPlus.put(code, 1);
                        break;
                    case "AMEX":
                        soharTotal[4]++;
                        if (soharAmex.containsKey(code))
                            soharAmex.put(code, soharAmex.get(code) + 1);
                        else
                            soharAmex.put(code, 1);
                        break;
                       
                    default:
                        if (soharOther.containsKey(code))
                            soharOther.put(code, soharOther.get(code) + 1);
                        else
                            soharOther.put(code, 1);
                    }
                    soharTotal[5]++;
                    if (soharCount.containsKey(code)) {
                        soharCount.put(code, soharCount.get(code) + 1);
                    } else {
                        soharCount.put(code, 1);
                    }
                }
               
                int soharSuccess = 0;
 
                System.out.println(soharCount);
 
                if (soharCount.get("000") != null && soharCount.get("00") != null) {
                    soharSuccess = soharCount.get("00") + soharCount.get("000");
                    soharCount.put("00", soharSuccess);
                    soharCount.remove("000");
                } else if (soharCount.get("000") == null && soharCount.get("00") == null) {
 
                } else {
                    if (soharCount.get("00") != null) {
                        soharSuccess = soharCount.get("00");
                    } else
                        soharSuccess = soharCount.get("000");
                    soharCount.put("00", soharSuccess);
                    soharCount.remove("000");
                }
               
                int soharDispute = 0;
               
                if (soharCount.get("000 Dispute Server") != null && soharCount.get("00 Dispute Server") != null) {
                    soharDispute = soharCount.get("00 Dispute Server") + soharCount.get("000 Dispute Server");
                    soharCount.put("00 Dispute Server", soharDispute);
                    soharCount.remove("000 Dispute Server");
                } else if (soharCount.get("000 Dispute Server") == null
                        && soharCount.get("00 Dispute Server") == null) {
 
                } else {
                    if (soharCount.get("00 Dispute Server") != null) {
                        soharDispute = soharCount.get("00 Dispute Server");
                    } else if (soharCount.get("000 Dispute Server") != null)
                        soharDispute = soharCount.get("000 Dispute Server");
                    soharCount.put("00 Dispute Server", soharDispute);
                    soharCount.remove("000 Dispute Server");
                }
               
                String soharSettlementStatus = "";
                HSSFRow soharHeaderRow = sheet.createRow(sheet.getLastRowNum() + 2);
                HSSFCell soharHeaderCell = soharHeaderRow.createCell(0);
                soharHeaderCell.setCellValue("Transactions Detail report - Bank Sohar");
                sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 7));
                soharHeaderCell.setCellStyle(style);
                text.append("<br>");
                text.append(
                        "<table border='1'  border = '1 px solid black' bordercolor='BLACK' style='border-collapse:collapse; font-family:Calibri;'>");
                text.append("<tr  border = '1 px solid black'>");
                text.append("<td colspan='8'  border = '1 px solid black'style ='text-align : center'>");
                text.append("<b>Transactions Detail report - Sohar</b>");
                text.append("</td>");
                text.append("</tr>");
                sheet.autoSizeColumn(0);
                soharHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
 
                // Makarand code for new output Sheet
                String bankBSH = "Financial Institution ID : BSH - Bank Sohar";
                String tranDateBSH = "Transaction Date : " + getCurrentOmanTime(); // append
                                                                                    // Oman
                                                                                    // Time
                String runDateBSH = "Run Date/Time : " + java.time.LocalDate.now() + " " + getCurrentOmanTime()
                        + " Asia/Muscat"; // Append IST Date and Oman Time
 
                for (int i = 1; i < 4; i++) {
                    soharHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    HSSFCell bankCell = soharHeaderRow.createCell(0);
                    sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 0, 7));
                    bankCell.setCellStyle(style1);
                    if (i == 1) {
 
                        bankCell.setCellValue(bankBSH);
                        text.append("<tr>");
                        text.append("<td colspan='8'   border = '1 px solid black'>");
                        text.append("<b>" + bankBSH + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                    if (i == 2) {
 
                        bankCell.setCellValue(tranDateBSH);
                        text.append("<tr>");
                        text.append("<td colspan='8'   border = '1 px solid black'>");
                        text.append("<b>" + tranDateBSH + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                    if (i == 3) {
 
                        bankCell.setCellValue(runDateBSH);
                        text.append("<tr>");
                        text.append("<td colspan='8'   border = '1 px solid black'>");
                        text.append("<b>" + runDateBSH + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                }
 
                /*
                 * Original code for (int i = 2; i <= 5; i++) { soharHeaderRow =
                 * sheet.createRow(sheet.getLastRowNum() + 1); HSSFCell bankCell
                 * = soharHeaderRow.createCell(0); sheet.addMergedRegion(new
                 * CellRangeAddress(sheet.getLastRowNum(),
                 * sheet.getLastRowNum(), 0, 6)); bankCell.setCellStyle(style1);
                 * bankCell.setCellValue(soharSheeet.getRow(i).getCell(0).
                 * toString()); text.append("<tr>"); text.append(
                 * "<td colspan='7'   border = '1 px solid black'>");
                 * text.append("<b>" +
                 * soharSheeet.getRow(i).getCell(0).toString() + "</b>");
                 * text.append("</td>"); text.append("</tr>"); }
                 *
                 */
 
                HSSFRow soharRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                HSSFCell soharCellTotal = soharRowTotal.createCell(0);
                soharCellTotal.setCellStyle(style);
                sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 1));
                soharCellTotal.setCellValue("Total Transactions");
                text.append("<tr>");
                text.append("<td border = '1 px solid black' colspan='2' rowspan='2'style ='text-align : center'>");
                text.append("<b>Total Transactions</b>");
                text.append("</td>");
                soharCellTotal = soharRowTotal.createCell(2);
                soharCellTotal.setCellStyle(style);
                soharCellTotal.setCellValue("VISA");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISA</b>");
                text.append("</td>");
                soharCellTotal = soharRowTotal.createCell(3);
                soharCellTotal.setCellStyle(style);
                soharCellTotal.setCellValue("MASTER CARD");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Master Card</b>");
                text.append("</td>");
                soharCellTotal = soharRowTotal.createCell(4);
                soharCellTotal.setCellStyle(style);
                soharCellTotal.setCellValue("HPS SOHAR");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>HPS Sohar</b>");
                text.append("</td>");
                soharCellTotal = soharRowTotal.createCell(5);
                soharCellTotal.setCellStyle(style);
                soharCellTotal.setCellValue("VISION PLUS");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISION PLUS</b>");
                text.append("</td>");
                soharCellTotal = soharRowTotal.createCell(6);
                soharCellTotal.setCellStyle(style);
                soharCellTotal.setCellValue("AMEX");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>AMEX</b>");
                text.append("</td>");
                soharCellTotal = soharRowTotal.createCell(7);
                soharCellTotal.setCellStyle(style);
                soharCellTotal.setCellValue("GRAND TOTAL");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>GRAND TOTAL</b>");
                text.append("</td>");
                text.append("</tr>");
                soharRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
               
                text.append("<tr>");
               
                for (int i = 0; i < soharTotal.length; i++) {
                   
                    text.append("<td  border = '1 px solid black'>");
                    text.append(soharTotal[i]);
                    text.append("</td>");
                   
                    int j = i + 2;
                   
                    soharCellTotal = soharRowTotal.createCell(j);
                    soharCellTotal.setCellStyle(style2);
                    soharCellTotal.setCellValue(soharTotal[i]);
                }
 
                text.append("</tr>");
               
                HSSFRow soharRowBlank = sheet.createRow(sheet.getLastRowNum() + 1);
                HSSFCell soharCellHeader = soharRowBlank.createCell(0);
                soharCellHeader.setCellStyle(style);
                soharCellHeader.setCellValue("Response code");
                text.append("<tr>");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Response code</b>");
                text.append("</td>");
                soharCellHeader = soharRowBlank.createCell(1);
                soharCellHeader.setCellStyle(style);
                soharCellHeader.setCellValue("Response code description");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Response code description</b>");
                text.append("</td>");
                soharCellHeader = soharRowBlank.createCell(2);
                soharCellHeader.setCellStyle(style);
                soharCellHeader.setCellValue("VISA");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISA</b>");
                text.append("</td>");
                soharCellHeader = soharRowBlank.createCell(3);
                soharCellHeader.setCellStyle(style);
                soharCellHeader.setCellValue("MASTERCARD");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>MASTERCARD</b>");
                text.append("</td>");
                soharCellHeader = soharRowBlank.createCell(4);
                soharCellHeader.setCellStyle(style);
                soharCellHeader.setCellValue("HPS SOHAR");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>HPS Sohar</b>");
                text.append("</td>");
                soharCellHeader = soharRowBlank.createCell(5);
                soharCellHeader.setCellStyle(style);
                soharCellHeader.setCellValue("VISIONPLUS");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISIONPLUS</b>");
                text.append("</td>");
                soharCellHeader = soharRowBlank.createCell(6);
                soharCellHeader.setCellStyle(style);
                soharCellHeader.setCellValue("AMEX");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>AMEX</b>");
                text.append("</td>");
                soharCellHeader = soharRowBlank.createCell(7);
                soharCellHeader.setCellStyle(style);
                soharCellHeader.setCellValue("Total count");
                soharCellHeader.setCellStyle(style);
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Total count</b>");
                text.append("</td>");
                text.append("</tr>");
               
                sheet.autoSizeColumn(0);
                sheet.autoSizeColumn(1);
                sheet.autoSizeColumn(2);
                sheet.autoSizeColumn(3);
                sheet.autoSizeColumn(4);
                sheet.autoSizeColumn(5);
                sheet.autoSizeColumn(6);
                sheet.autoSizeColumn(7);
               
                int soharLength = sheet.getLastRowNum() + 1;
               
                Set<String> soharKeys = soharCount.keySet();
                // System.out.println("******************** SET Sohar keys :
                // "+soharKeys);
                Iterator<String> soharIt = soharKeys.iterator();
               
                for (int i = 0; i < soharCount.size(); i++) {
                    text.append("<tr>");
 
                    String temp = soharIt.next();
                    // System.out.println("**************** Temp Variable :
                    // "+temp);
                    HSSFRow dataRow = sheet.createRow(soharLength++);
                    dataRow.setRowStyle(style1);
                    HSSFCell dataCell = dataRow.createCell(0);
                    dataCell.setCellStyle(style2);
                    dataCell.setCellValue(temp);
                    text.append("<td border = '1 px solid black'>");
                    text.append(temp);
                    text.append("</td>");
                                       
                    if (temp.length() > 4) {
                        soharSettlementStatus = temp.substring(3, temp.length());
                        System.out.println(soharSettlementStatus);
                    }
                   
                    dataCell = dataRow.createCell(1);
                    dataCell.setCellStyle(style2);
                   
                    if (prop.getProperty(temp) != null) {
                        dataCell.setCellValue(prop.getProperty(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append(prop.getProperty(temp));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue("unknown response code/ unknown status");
                        text.append("<td border = '1 px solid black'>");
                        text.append("unknown response code/ unknown status");
                        text.append("</td>");
                    }
                   
                    dataCell = dataRow.createCell(2);
                    dataCell.setCellStyle(style2);
                    if (soharVisa.get(temp) != null) {
                        text.append("<td border = '1 px solid black'>");
                        text.append((soharVisa.get(temp)));
                        text.append("</td>");
                        dataCell.setCellValue(soharVisa.get(temp));
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                   
                    dataCell = dataRow.createCell(3);
                    dataCell.setCellStyle(style2);
                    if (soharMastercard.get(temp) != null) {
                        text.append("<td border = '1 px solid black'>");
                        text.append((soharMastercard.get(temp)));
                        text.append("</td>");
                        dataCell.setCellValue(soharMastercard.get(temp));
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                   
                    dataCell = dataRow.createCell(4);
                    dataCell.setCellStyle(style2);
                    if (temp.equals("00")) {
                        temp = "000";
                        if (soharHpsSohar.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((soharHpsSohar.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(soharHpsSohar.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00";
                    } else if (temp.equals("00 Dispute Server")) {
                        temp = "000 Dispute Server";
                        if (soharHpsSohar.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((soharHpsSohar.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(soharHpsSohar.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00 Dispute Server";
                    } else {
                        if (soharHpsSohar.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((soharHpsSohar.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(soharHpsSohar.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                    }
                   
                    dataCell = dataRow.createCell(5);
                    dataCell.setCellStyle(style2);
                    if (soharVisionPlus.get(temp) != null) {
                        text.append("<td border = '1 px solid black'>");
                        text.append((soharVisionPlus.get(temp)));
                        text.append("</td>");
                        dataCell.setCellValue(soharVisionPlus.get(temp));
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                   
                    dataCell = dataRow.createCell(6);
                    dataCell.setCellStyle(style2);
                   
                    if (temp.equals("00")) {
                        temp = "000";
                       
                        if (soharAmex.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((soharAmex.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(soharAmex.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00";
                    } else if (temp.equals("00 Dispute Server")) {
                        temp = "000 Dispute Server";
                        if (soharAmex.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((soharAmex.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(soharAmex.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00 Dispute Server";
                    } else {
                        if (soharAmex.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((soharAmex.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(soharAmex.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                    }
                   
                                   
                    dataCell = dataRow.createCell(7);
                    dataCell.setCellStyle(style2);
                    dataCell.setCellValue(soharCount.get(temp));
                    text.append("<td border = '1 px solid black'>");
                    text.append(soharCount.get(temp));
                    text.append("</td>");
                    text.append("</tr>");
 
                }
 
                if (tester == 0) {
                    rcText.append("<tr  border = '1 px solid black'>");
                    rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                    rcText.append("<b>-- Null --</b>" + "</td></tr>");
                }
 
                tester = 0;
                rcText.append("</table>");
 
                text.append("</table>");
                for (int k = 0; k < 8; k++) {
                    sheet.autoSizeColumn(k);
                }
                tot = sheet.getLastRowNum();
                AUBInitial = tot;
                for (int i = soharInitial + 1; i < tot; i++) {
                    int j = i + 1;
                    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 7), sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 7), sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 7), sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 7), sheet);
                }
                for (int i = 0; i < 6; i++) {
                    int j = i + 1;
                    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(soharInitial + 5, tot, i, j),
                            sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(soharInitial + 5, tot, i, j),
                            sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(soharInitial + 5, tot, i, j),
                            sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(soharInitial + 5, tot, i, j), sheet);
 
                }
 
                /*
                 * To maintain backup of report renaming previous report file
                 * with extention of timstamp
                 */
 
                File oldInfileName = new File(prop.getProperty("ManualsoharInFile") + ".xlsx");
                File newInfileName = new File(prop.getProperty("ManualsoharInFile") + timeStamp + ".xlsx");
 
                oldInfileName.renameTo(newInfileName);
                System.out.println("Rename completed successfully");
 
                prop.setProperty("bshTotalCurr", Integer.toString(soharTotal[5]));
 
               
            } else {
                System.out.println("BSH Infile not present");
            }
        } catch (Exception e) {
 
            System.out.println("Exception in BSH report output...");
            fin.close();
            fo.close();
            System.out.println(e);
 
        }
 
        // ********************************************************************************************************
 
        FileOutputStream fos = new FileOutputStream(new File(prop.getProperty("output") + timeStamp + ".xls"));
        book.write(fos);
        fos.flush();
        fos.close();
        System.out.println("written ");
 
        text.append("</tr>");
        text.append("</table>");
 
  //  SendEmail email = new SendEmail();
    //  email.sendReportMail(text);
 
       
        prop.setProperty("aboTotalPrev", prop.getProperty("aboTotalCurr"));
        prop.setProperty("afsTotalPrev", prop.getProperty("afsTotalCurr"));
        prop.setProperty("aubTotalPrev", prop.getProperty("aubTotalCurr"));
        prop.setProperty("bshTotalPrev", prop.getProperty("bshTotalCurr"));
 
        prop.store(fo, null);
        fo.close();
        fin.close();
 
       
    }
   
    /*
     *
     * NORMAL REPORT OUTPUT EXECUTION END
     *
     */
 
    public void transCountOman() throws IOException {
        // TODO Auto-generated method stub
 
        String timeStamp = new SimpleDateFormat("ddMM_HHmm").format(Calendar.getInstance().getTime());
        String timeStamp1 = new SimpleDateFormat("ddMM").format(Calendar.getInstance().getTime());
 
        System.out.println("Timestamp 1 :- " + timeStamp1);
 
        initialize();
 
        // Critical RC Null chcking flag
        int tester = 0;
 
        StringBuilder text = new StringBuilder();
 
        StringBuilder rcText = new StringBuilder();
 
        TreeMap<String, Integer> count = new TreeMap<>();
        TreeMap<String, Integer> visa = new TreeMap<>();
        TreeMap<String, Integer> mastercard = new TreeMap<>();
        TreeMap<String, Integer> omannet = new TreeMap<>();
        TreeMap<String, Integer> visionPlus = new TreeMap<>();
        TreeMap<String, Integer> other = new TreeMap<>();
        int tot;
        int bahrainInitial = 0, soharInitial = 0, AUBInitial = 0;
        int[] total = new int[5];
 
        /* Compare RC chart preparation */
 
        rcText.append("<table border='1' bordercolor='BLACK' style='border-collapse:collapse; font-family:Calibri;'>");
        rcText.append("<tr  border = '1 px solid black'>");
 
        rcText.append("<td colspan='1'  border = '1 px solid black'style ='text-align : center'>");
        rcText.append("<b>Response Code</b>" + "</td>");
        rcText.append("<td colspan='1'  border = '1 px solid black'style ='text-align : center'>");
        rcText.append("<b>Previous Count</b>" + "</td>");
        rcText.append("<td colspan='1'  border = '1 px solid black'style ='text-align : center'>");
        rcText.append("<b>Current Count</b>" + "</td>");
 
        rcText.append("</tr>");
 
        /* End of this section */
 
        HSSFWorkbook book = new HSSFWorkbook();
        HSSFSheet sheet = book.createSheet("transactions");
 
        HSSFCellStyle style = book.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont font = book.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        font.setFontName("Calibri");
        style.setFont(font);
        HSSFCellStyle style1 = book.createCellStyle();
        style1.setAlignment(HorizontalAlignment.LEFT);
        HSSFFont font1 = book.createFont();
        font1.setFontHeightInPoints((short) 12);
        font1.setBold(true);
        font1.setFontName("Calibri");
        style1.setFont(font1);
        HSSFCellStyle style2 = book.createCellStyle();
        HSSFFont font2 = book.createFont();
        font2.setFontHeightInPoints((short) 12);
        // font2.setBoldweight((short) 1000000);
        font2.setFontName("Calibri");
        style2.setAlignment(HorizontalAlignment.LEFT);
        style2.setFont(font2);
        try {
            if (new File(prop.getProperty("omanInFile") + ".xlsx").exists()) {
                FileInputStream updated = new FileInputStream(prop.getProperty("omanInFile") + ".xlsx");
                XSSFWorkbook workbook = new XSSFWorkbook(updated);
                // FileOutputStream fos = new FileOutputStream(new
                // File(prop.getProperty("omanBackup")));
                // workbook.write(fos);
                XSSFSheet sheeet = workbook.getSheet(workbook.getSheetName(0));
 
                rcText.append("<tr  border = '1 px solid black'>");
                rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                rcText.append("<b>========= ABO ===========</b>" + "</td></tr>");
                for (int i = 8; i <= sheeet.getLastRowNum() - 2; i++) {
                    XSSFRow row = sheeet.getRow(i);
                    XSSFCell cell = row.getCell(18); // Resp Code
                    String code = cell.getStringCellValue();
                    XSSFCell cell2 = row.getCell(17); // Interchange
                    String interchange = cell2.getStringCellValue();
                    cell2 = row.getCell(19); // Settlement stat ; 20 = Tran stat
                    String status = cell2.getStringCellValue();
 
                    if ((code.equals("") || code.equals(null) || code.equals(" ") || code.equals("0"))) {
                        if (row.getCell(20).getStringCellValue().equalsIgnoreCase("In progress")) {
                            code = new String(row.getCell(20).getStringCellValue());
                        } else {
                            code = row.getCell(20).getStringCellValue();
                        }
                    } else if (row.getCell(20).getStringCellValue().equalsIgnoreCase("Timeout")) {
 
                        code = code + " " + new String(row.getCell(20).getStringCellValue());
 
                    } else if (!(status.equals("Not initiated") || status.equals("Settled"))) {
                        code = code + " " + status;
                    }
 
                    /*
                     * Makarand added code if
                     * (row.getCell(19).getStringCellValue().equalsIgnoreCase(
                     * "Timeout")) { code = code + " " + new
                     * String(row.getCell(19).getStringCellValue()); }
                     */
                    switch (interchange) {
                    case "OMANNET":
                        total[2]++;
                        if (omannet.containsKey(code))
                            omannet.put(code, omannet.get(code) + 1);
                        else
                            omannet.put(code, 1);
                        break;
                    case "MASTERCARD":
                        total[1]++;
                        if (mastercard.containsKey(code))
                            mastercard.put(code, mastercard.get(code) + 1);
                        else
                            mastercard.put(code, 1);
                        break;
                    case "VISAAHB":
                        total[0]++;
                        if (visa.containsKey(code))
                            visa.put(code, visa.get(code) + 1);
                        else
                            visa.put(code, 1);
                        break;
                    case "VISIONPLUSAHB":
                        total[3]++;
                        if (visionPlus.containsKey(code))
                            visionPlus.put(code, visionPlus.get(code) + 1);
                        else
                            visionPlus.put(code, 1);
                        break;
                    default:
                        if (other.containsKey(code))
                            other.put(code, other.get(code) + 1);
                        else
                            other.put(code, 1);
                    }
                    total[4]++;
                    if (count.containsKey(code)) {
                        count.put(code, count.get(code) + 1);
                    } else {
                        count.put(code, 1);
                    }
                }
                int success;
                if (count.get("000") != null && count.get("00") != null) {
                    success = count.get("00") + count.get("000");
                    count.put("00", success);
                    count.remove("000");
                } else if (count.get("000") == null && count.get("00") == null) {
 
                } else {
                    if (count.get("00") != null) {
                        success = count.get("00");
                    } else
                        success = count.get("000");
                    count.put("00", success);
                    count.remove("000");
                }
                int dispute = 0;
                if (count.get("000 Dispute Server") != null && count.get("00 Dispute Server") != null) {
                    dispute = count.get("00 Dispute Server") + count.get("000 Dispute Server");
                    count.put("00 Dispute Server", dispute);
                    count.remove("000 Dispute Server");
                } else if (count.get("000 Dispute Server") == null && count.get("00 Dispute Server") == null) {
 
                } else {
                    if (count.get("00 Dispute Server") != null) {
                        dispute = count.get("00 Dispute Server");
                    } else if (count.get("000 Dispute Server") != null)
                        dispute = count.get("000 Dispute Server");
                    count.put("00 Dispute Server", dispute);
                    count.remove("000 Dispute Server");
                }
                HSSFRow headerRow = sheet.createRow(0);
                HSSFCell headerCell = headerRow.createCell(0);
                headerCell.setCellValue("Transactions Detail report - Oman");
 
                text.append(
                        "<table border='1' bordercolor='BLACK' style='border-collapse:collapse; font-family:Calibri;'>");
                text.append("<tr  border = '1 px solid black'>");
                text.append("<td colspan='7'  border = '1 px solid black'style ='text-align : center'>");
                text.append("<b>Transactions Detail report - Oman</b>");
                text.append("</td>");
                text.append("</tr>");
                sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 6));
                headerCell.setCellStyle(style);
                sheet.autoSizeColumn(0);
                headerRow = sheet.createRow(sheet.getLastRowNum() + 1);
 
                // Makarand code for new output Sheet
                String bankABO = "Financial Institution ID : AHB - Ahli Bank of Oman";
                String tranDateABO = "Transaction Date : " + getCurrentOmanTime(); // append
                                                                                    // Oman
                                                                                    // Time
                String runDateABO = "Run Date/Time : " + java.time.LocalDate.now() + " " + getCurrentOmanTime()
                        + " Asia/Muscat"; // Append IST Date and Oman Time
                for (int i = 1; i < 4; i++) {
                    headerRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    HSSFCell bankCell = headerRow.createCell(0);
                    sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 0, 6));
                    bankCell.setCellStyle(style1);
                    if (i == 1) {
 
                        bankCell.setCellValue(bankABO);
                        text.append("<tr>");
                        text.append("<td colspan='7'   border = '1 px solid black'>");
                        text.append("<b>" + bankABO + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                    if (i == 2) {
 
                        bankCell.setCellValue(tranDateABO);
                        text.append("<tr>");
                        text.append("<td colspan='7'   border = '1 px solid black'>");
                        text.append("<b>" + tranDateABO + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
 
                    if (i == 3) {
 
                        bankCell.setCellValue(runDateABO);
                        text.append("<tr>");
                        text.append("<td colspan='7'   border = '1 px solid black'>");
                        text.append("<b>" + runDateABO + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }
                }
                HSSFRow rowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                text.append("<tr>");
                HSSFCell cellTotal = rowTotal.createCell(0);
 
                cellTotal.setCellStyle(style);
                sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 1));
                cellTotal.setCellValue("Total Transactions>");
                text.append("<td border = '1 px solid black' colspan='2' rowspan='2'style ='text-align : center'>");
                text.append("<b>Total Transactions</b>");
                text.append("</td>");
                cellTotal = rowTotal.createCell(2);
                cellTotal.setCellStyle(style);
                cellTotal.setCellValue("VISA");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISA</b>");
                text.append("</td>");
                cellTotal = rowTotal.createCell(3);
                cellTotal.setCellStyle(style);
                cellTotal.setCellValue("MASTER CARD");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Master Card</b>");
                text.append("</td>");
                cellTotal = rowTotal.createCell(4);
                cellTotal.setCellStyle(style);
                cellTotal.setCellValue("OMANNET");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>OMANNET</b>");
                text.append("</td>");
                cellTotal = rowTotal.createCell(5);
                cellTotal.setCellStyle(style);
                cellTotal.setCellValue("VISION PLUS");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISION PLUS</b>");
                text.append("</td>");
                cellTotal = rowTotal.createCell(6);
                cellTotal.setCellStyle(style);
                cellTotal.setCellValue("GRAND TOTAL");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>GRAND TOTAL</b>");
                text.append("</td>");
                text.append("</tr>");
                rowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                text.append("<tr>");
 
                for (int i = 0; i < total.length; i++) {
                    int j = i + 2;
                    cellTotal = rowTotal.createCell(j);
                    cellTotal.setCellStyle(style2);
                    cellTotal.setCellValue(total[i]);
                    text.append("<td  border = '1 px solid black'>");
                    text.append(total[i]);
                    text.append("</td>");
                }
                text.append("</tr>");
                HSSFRow rowBlank = sheet.createRow(sheet.getLastRowNum() + 1);
                text.append("<tr>");
 
                HSSFCell cellHeader = rowBlank.createCell(0);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("Response code");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Response code</b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(1);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("Response code description");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Response code description/<b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(2);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("VISA");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISA</b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(3);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("MASTERCARD");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>MASTERCARD</b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(4);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("OMANNET");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>OMANNET</b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(5);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("VISIONPLUS");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISIONPLUS</b>");
                text.append("</td>");
                cellHeader = rowBlank.createCell(6);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue("Total count");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Total count</b>");
                text.append("</td>");
                text.append("</tr>");
                cellHeader.setCellStyle(style);
                sheet.autoSizeColumn(0);
                sheet.autoSizeColumn(1);
                sheet.autoSizeColumn(2);
                sheet.autoSizeColumn(3);
                sheet.autoSizeColumn(4);
                sheet.autoSizeColumn(5);
 
                System.out.println("Oman Count : " + count);
 
                int length = sheet.getLastRowNum() + 1;
                Set<String> keys = count.keySet();
                // System.out.println(keys);
                Iterator<String> it = keys.iterator();
                for (int i = 0; i < count.size(); i++) {
                    text.append("<tr>");
                    String temp = it.next();
                    HSSFRow dataRow = sheet.createRow(length++);
                    dataRow.setRowStyle(style);
                    HSSFCell dataCell = dataRow.createCell(0);
                    dataCell.setCellStyle(style2);
                    dataCell.setCellValue(temp);
                    text.append("<td border = '1 px solid black'>");
                    text.append(temp);
                    text.append("</td>");
                    if (temp.length() > 4) {
                    }
                    dataCell = dataRow.createCell(1);
                    dataCell.setCellStyle(style2);
                    // if (flag) {
                    // dataCell.setCellValue(settlementStatus);
                    //
                    // } else {
                    if (prop.getProperty(temp) != null) {
                        dataCell.setCellValue(prop.getProperty(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append(prop.getProperty(temp));
                        text.append("</td>");
                    }
 
                    else {
                        dataCell.setCellValue("unknown response code/ unknown status");
                        text.append("<td border = '1 px solid black'>");
                        text.append("unknown response code/ unknown status");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(2);
                    dataCell.setCellStyle(style2);
                    if (visa.get(temp) != null) {
                        text.append("<td border = '1 px solid black'>");
                        text.append((visa.get(temp)));
                        text.append("</td>");
                        dataCell.setCellValue(visa.get(temp));
                    }
 
                    else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
 
                    dataCell = dataRow.createCell(3);
                    dataCell.setCellStyle(style2);
                    if (mastercard.get(temp) != null) {
                        text.append("<td border = '1 px solid black'>");
                        text.append((mastercard.get(temp)));
                        text.append("</td>");
                        dataCell.setCellValue(mastercard.get(temp));
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(4);
                    dataCell.setCellStyle(style2);
                    if (temp.equals("00")) {
                        temp = "000";
                        if (omannet.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((omannet.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(omannet.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00";
                    } else if (temp.equals("00 Dispute Server")) {
                        temp = "000 Dispute Server";
                        if (omannet.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((omannet.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(omannet.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00 Dispute Server";
                    } else {
                        if (omannet.get(temp) != null) {
                            text.append("<td border = '1 px solid black'>");
                            text.append((omannet.get(temp)));
                            text.append("</td>");
                            dataCell.setCellValue(omannet.get(temp));
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                    }
                    dataCell = dataRow.createCell(5);
                    dataCell.setCellStyle(style2);
                    if (visionPlus.get(temp) != null) {
                        dataCell.setCellValue(visionPlus.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((visionPlus.get(temp)));
                        text.append("</td>");
                    }
 
                    else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(6);
                    dataCell.setCellStyle(style2);
                    dataCell.setCellValue(count.get(temp));
                    text.append("<td border = '1 px solid black'>");
                    text.append(count.get(temp));
                    text.append("</td>");
                    text.append("</tr>");
 
                }
                if (tester == 0) {
                    rcText.append("<tr  border = '1 px solid black'>");
                    rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                    rcText.append("<b>-- Null --</b>" + "</td></tr>");
                }
 
                tester = 0;
 
                text.append("</table>");
                tot = sheet.getLastRowNum();
                bahrainInitial = tot + 1;
                for (int i = 0; i < tot; i++) {
                    int j = i + 1;
                    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                }
                for (int i = 1; i < 6; i++) {
                    int j = i + 1;
                    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(6, tot, i, j), sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(6, tot, i, j), sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(6, tot, i, j), sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(6, tot, i, j), sheet);
 
                }
                for (int k = 0; k < 21; k++) {
                    sheet.autoSizeColumn(k);
                }
 
                /*
                 * To maintain backup of report renaming previous report file
                 * with extention of timstamp
                 */
 
                File oldInfileName = new File(prop.getProperty("omanInFile") + ".xls");
                File newInfileName = new File(prop.getProperty("omanInFile") + timeStamp + ".xls");
 
                oldInfileName.renameTo(newInfileName);
                System.out.println("Rename completed successfully");
 
            }
        } catch (Exception e) {
 
            System.out.println("Exception in ABO report output...");
            fin.close();
            fo.close();
            System.out.println(e);
        }
       
       
        // ****************************** Sohar Environment
                // ************************************************
                try {
                    if (new File(prop.getProperty("soharInFile") + ".xlsx").exists()) {
                        TreeMap<String, Integer> soharCount = new TreeMap<>();
                        TreeMap<String, Integer> soharVisa = new TreeMap<>();
                        TreeMap<String, Integer> soharMastercard = new TreeMap<>();
                        TreeMap<String, Integer> soharHpsSohar = new TreeMap<>();
                        TreeMap<String, Integer> soharVisionPlus = new TreeMap<>();
                        TreeMap<String, Integer> soharOther = new TreeMap<>();
                        int[] soharTotal = new int[5];
 
                        FileInputStream soharUpdated = new FileInputStream(prop.getProperty("soharInFile") + ".xlsx");
                        XSSFWorkbook soharWorkbook = new XSSFWorkbook(soharUpdated);
                        // FileOutputStream fos = new FileOutputStream(new
                        // File(prop.getProperty("soharBackup")));
                        // soharWorkbook.write(fos);
                        XSSFSheet soharSheeet = soharWorkbook.getSheet(soharWorkbook.getSheetName(0));
 
                        rcText.append("<tr  border = '1 px solid black'>");
                        rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                        rcText.append("<b>========= BSH ===========</b>" + "</td></tr>");
 
                        for (int i = 8; i <= soharSheeet.getLastRowNum() - 2; i++) { // initial
                                                                                        // i=8
                            XSSFRow soharRow = soharSheeet.getRow(i);
                            /*
                             * System.out.println(
                             * "**********************************************");
                             * System.out.print("BSH cell 17"
                             * +soharRow.getCell(17).getStringCellValue());
                             * System.out.print("BSH cell 18"
                             * +soharRow.getCell(18).getStringCellValue());
                             * System.out.print("BSH cell 19"
                             * +soharRow.getCell(19).getStringCellValue());
                             * System.out.print("BSH cell 20"
                             * +soharRow.getCell(20).getStringCellValue());
                             * System.out.println(
                             * "**********************************************");
                             */
 
                            XSSFCell cell = soharRow.getCell(18); // Resp Code
                            String code = cell.getStringCellValue();
                            // System.out.println("******************************** CODE
                            // before : "+code);
                            XSSFCell cell2 = soharRow.getCell(17); // Interchange
                            String interchange = cell2.getStringCellValue();
                            cell2 = soharRow.getCell(19); // Settlement Status ; col 19
                                                            // = Tran status
                            String status = cell2.getStringCellValue();
 
                            if ((code.equals("") || code.equals(null) || code.equals(" ") || code.equals("0"))) {
                                if (soharRow.getCell(20).getStringCellValue().equalsIgnoreCase("In progress")) {
                                    code = new String("In progress");
                                } else {
                                    code = soharRow.getCell(20).getStringCellValue();
                                }
                            } else if (soharRow.getCell(20).getStringCellValue().equalsIgnoreCase("Timeout")) {
 
                                code = code + " " + new String(soharRow.getCell(20).getStringCellValue());
 
                            } else if (!(status.equals("Not initiated") || status.equals("Settled"))) {
                                code = code + " " + status;
                            }
 
                            switch (interchange) {
                            case "HPS SOHAR":
                                soharTotal[2]++;
                                if (soharHpsSohar.containsKey(code))
                                    soharHpsSohar.put(code, soharHpsSohar.get(code) + 1);
                                else
                                    soharHpsSohar.put(code, 1);
                                break;
                            case "MASTERCARD":
                                soharTotal[1]++;
                                if (soharMastercard.containsKey(code))
                                    soharMastercard.put(code, soharMastercard.get(code) + 1);
                                else
                                    soharMastercard.put(code, 1);
                                break;
                            case "VISA":
                                soharTotal[0]++;
                                if (soharVisa.containsKey(code))
                                    soharVisa.put(code, soharVisa.get(code) + 1);
                                else
                                    soharVisa.put(code, 1);
                                break;
                            case "VISION PLUS":
                                soharTotal[3]++;
                                if (soharVisionPlus.containsKey(code))
                                    soharVisionPlus.put(code, soharVisionPlus.get(code) + 1);
                                else
                                    soharVisionPlus.put(code, 1);
                                break;
                            default:
                                if (soharOther.containsKey(code))
                                    soharOther.put(code, soharOther.get(code) + 1);
                                else
                                    soharOther.put(code, 1);
                            }
                            soharTotal[4]++;
                            if (soharCount.containsKey(code)) {
                                soharCount.put(code, soharCount.get(code) + 1);
                            } else {
                                soharCount.put(code, 1);
                            }
                        }
                        int soharSuccess = 0;
 
                        System.out.println(soharCount);
 
                        if (soharCount.get("000") != null && soharCount.get("00") != null) {
                            soharSuccess = soharCount.get("00") + soharCount.get("000");
                            soharCount.put("00", soharSuccess);
                            soharCount.remove("000");
                        } else if (soharCount.get("000") == null && soharCount.get("00") == null) {
 
                        } else {
                            if (soharCount.get("00") != null) {
                                soharSuccess = soharCount.get("00");
                            } else
                                soharSuccess = soharCount.get("000");
                            soharCount.put("00", soharSuccess);
                            soharCount.remove("000");
                        }
                        int soharDispute = 0;
                        if (soharCount.get("000 Dispute Server") != null && soharCount.get("00 Dispute Server") != null) {
                            soharDispute = soharCount.get("00 Dispute Server") + soharCount.get("000 Dispute Server");
                            soharCount.put("00 Dispute Server", soharDispute);
                            soharCount.remove("000 Dispute Server");
                        } else if (soharCount.get("000 Dispute Server") == null
                                && soharCount.get("00 Dispute Server") == null) {
 
                        } else {
                            if (soharCount.get("00 Dispute Server") != null) {
                                soharDispute = soharCount.get("00 Dispute Server");
                            } else if (soharCount.get("000 Dispute Server") != null)
                                soharDispute = soharCount.get("000 Dispute Server");
                            soharCount.put("00 Dispute Server", soharDispute);
                            soharCount.remove("000 Dispute Server");
                        }
                        String soharSettlementStatus = "";
                        HSSFRow soharHeaderRow = sheet.createRow(sheet.getLastRowNum() + 2);
                        HSSFCell soharHeaderCell = soharHeaderRow.createCell(0);
                        soharHeaderCell.setCellValue("Transactions Detail report - Bank Sohar");
                        sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 6));
                        soharHeaderCell.setCellStyle(style);
                        text.append("<br>");
                        text.append(
                                "<table border='1'  border = '1 px solid black' bordercolor='BLACK' style='border-collapse:collapse; font-family:Calibri;'>");
                        text.append("<tr  border = '1 px solid black'>");
                        text.append("<td colspan='7'  border = '1 px solid black'style ='text-align : center'>");
                        text.append("<b>Transactions Detail report - Sohar</b>");
                        text.append("</td>");
                        text.append("</tr>");
                        sheet.autoSizeColumn(0);
                        soharHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
 
                        // Makarand code for new output Sheet
                        String bankBSH = "Financial Institution ID : BSH - Bank Sohar";
                        String tranDateBSH = "Transaction Date : " + getCurrentOmanTime(); // append
                                                                                            // Oman
                                                                                            // Time
                        String runDateBSH = "Run Date/Time : " + java.time.LocalDate.now() + " " + getCurrentOmanTime()
                                + " Asia/Muscat"; // Append IST Date and Oman Time
 
                        for (int i = 1; i < 4; i++) {
                            soharHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
                            HSSFCell bankCell = soharHeaderRow.createCell(0);
                            sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 0, 6));
                            bankCell.setCellStyle(style1);
                            if (i == 1) {
 
                                bankCell.setCellValue(bankBSH);
                                text.append("<tr>");
                                text.append("<td colspan='7'   border = '1 px solid black'>");
                                text.append("<b>" + bankBSH + "</b>");
                                text.append("</td>");
                                text.append("</tr>");
                            }
 
                            if (i == 2) {
 
                                bankCell.setCellValue(tranDateBSH);
                                text.append("<tr>");
                                text.append("<td colspan='7'   border = '1 px solid black'>");
                                text.append("<b>" + tranDateBSH + "</b>");
                                text.append("</td>");
                                text.append("</tr>");
                            }
 
                            if (i == 3) {
 
                                bankCell.setCellValue(runDateBSH);
                                text.append("<tr>");
                                text.append("<td colspan='7'   border = '1 px solid black'>");
                                text.append("<b>" + runDateBSH + "</b>");
                                text.append("</td>");
                                text.append("</tr>");
                            }
 
                        }
 
                        /*
                         * Original code for (int i = 2; i <= 5; i++) { soharHeaderRow =
                         * sheet.createRow(sheet.getLastRowNum() + 1); HSSFCell bankCell
                         * = soharHeaderRow.createCell(0); sheet.addMergedRegion(new
                         * CellRangeAddress(sheet.getLastRowNum(),
                         * sheet.getLastRowNum(), 0, 6)); bankCell.setCellStyle(style1);
                         * bankCell.setCellValue(soharSheeet.getRow(i).getCell(0).
                         * toString()); text.append("<tr>"); text.append(
                         * "<td colspan='7'   border = '1 px solid black'>");
                         * text.append("<b>" +
                         * soharSheeet.getRow(i).getCell(0).toString() + "</b>");
                         * text.append("</td>"); text.append("</tr>"); }
                         *
                         */
 
                        HSSFRow soharRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                        HSSFCell soharCellTotal = soharRowTotal.createCell(0);
                        soharCellTotal.setCellStyle(style);
                        sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 1));
                        soharCellTotal.setCellValue("Total Transactions");
                        text.append("<tr>");
                        text.append("<td border = '1 px solid black' colspan='2' rowspan='2'style ='text-align : center'>");
                        text.append("<b>Total Transactions</b>");
                        text.append("</td>");
                        soharCellTotal = soharRowTotal.createCell(2);
                        soharCellTotal.setCellStyle(style);
                        soharCellTotal.setCellValue("VISA");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>VISA</b>");
                        text.append("</td>");
                        soharCellTotal = soharRowTotal.createCell(3);
                        soharCellTotal.setCellStyle(style);
                        soharCellTotal.setCellValue("MASTER CARD");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>Master Card</b>");
                        text.append("</td>");
                        soharCellTotal = soharRowTotal.createCell(4);
                        soharCellTotal.setCellStyle(style);
                        soharCellTotal.setCellValue("HPS SOHAR");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>HPS Sohar</b>");
                        text.append("</td>");
                        soharCellTotal = soharRowTotal.createCell(5);
                        soharCellTotal.setCellStyle(style);
                        soharCellTotal.setCellValue("VISION PLUS");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>VISION PLUS</b>");
                        text.append("</td>");
                        soharCellTotal = soharRowTotal.createCell(6);
                        soharCellTotal.setCellStyle(style);
                        soharCellTotal.setCellValue("GRAND TOTAL");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>GRAND TOTAL</b>");
                        text.append("</td>");
                        text.append("</tr>");
                        soharRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                        text.append("<tr>");
                        for (int i = 0; i < soharTotal.length; i++) {
                            text.append("<td  border = '1 px solid black'>");
                            text.append(soharTotal[i]);
                            text.append("</td>");
                            int j = i + 2;
                            soharCellTotal = soharRowTotal.createCell(j);
                            soharCellTotal.setCellStyle(style2);
                            soharCellTotal.setCellValue(soharTotal[i]);
                        }
 
                        text.append("</tr>");
                        HSSFRow soharRowBlank = sheet.createRow(sheet.getLastRowNum() + 1);
                        HSSFCell soharCellHeader = soharRowBlank.createCell(0);
                        soharCellHeader.setCellStyle(style);
                        soharCellHeader.setCellValue("Response code");
                        text.append("<tr>");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>Response code</b>");
                        text.append("</td>");
                        soharCellHeader = soharRowBlank.createCell(1);
                        soharCellHeader.setCellStyle(style);
                        soharCellHeader.setCellValue("Response code description");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>Response code description</b>");
                        text.append("</td>");
                        soharCellHeader = soharRowBlank.createCell(2);
                        soharCellHeader.setCellStyle(style);
                        soharCellHeader.setCellValue("VISA");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>VISA</b>");
                        text.append("</td>");
                        soharCellHeader = soharRowBlank.createCell(3);
                        soharCellHeader.setCellStyle(style);
                        soharCellHeader.setCellValue("MASTERCARD");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>MASTERCARD</b>");
                        text.append("</td>");
                        soharCellHeader = soharRowBlank.createCell(4);
                        soharCellHeader.setCellStyle(style);
                        soharCellHeader.setCellValue("HPS SOHAR");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>HPS Sohar</b>");
                        text.append("</td>");
                        soharCellHeader = soharRowBlank.createCell(5);
                        soharCellHeader.setCellStyle(style);
                        soharCellHeader.setCellValue("VISIONPLUS");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>VISIONPLUS</b>");
                        text.append("</td>");
                        soharCellHeader = soharRowBlank.createCell(6);
                        soharCellHeader.setCellStyle(style);
                        soharCellHeader.setCellValue("Total count");
                        soharCellHeader.setCellStyle(style);
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>Total count</b>");
                        text.append("</td>");
                        text.append("</tr>");
                        sheet.autoSizeColumn(0);
                        sheet.autoSizeColumn(1);
                        sheet.autoSizeColumn(2);
                        sheet.autoSizeColumn(3);
                        sheet.autoSizeColumn(4);
                        sheet.autoSizeColumn(5);
                        int soharLength = sheet.getLastRowNum() + 1;
                        Set<String> soharKeys = soharCount.keySet();
                        // System.out.println("******************** SET Sohar keys :
                        // "+soharKeys);
                        Iterator<String> soharIt = soharKeys.iterator();
                        for (int i = 0; i < soharCount.size(); i++) {
                            text.append("<tr>");
 
                            String temp = soharIt.next();
                            // System.out.println("**************** Temp Variable :
                            // "+temp);
                            HSSFRow dataRow = sheet.createRow(soharLength++);
                            dataRow.setRowStyle(style1);
                            HSSFCell dataCell = dataRow.createCell(0);
                            dataCell.setCellStyle(style2);
                            dataCell.setCellValue(temp);
                            text.append("<td border = '1 px solid black'>");
                            text.append(temp);
                            text.append("</td>");
                            if (temp.length() > 4) {
                                soharSettlementStatus = temp.substring(3, temp.length());
                                System.out.println(soharSettlementStatus);
                            }
                            dataCell = dataRow.createCell(1);
                            dataCell.setCellStyle(style2);
                            // if (flag) {
                            // dataCell.setCellValue(settlementStatus);
                            //
                            // } else {
                            if (prop.getProperty(temp) != null) {
                                dataCell.setCellValue(prop.getProperty(temp));
                                text.append("<td border = '1 px solid black'>");
                                text.append(prop.getProperty(temp));
                                text.append("</td>");
                            } else {
                                dataCell.setCellValue("unknown response code/ unknown status");
                                text.append("<td border = '1 px solid black'>");
                                text.append("unknown response code/ unknown status");
                                text.append("</td>");
                            }
                            // }
                            dataCell = dataRow.createCell(2);
                            dataCell.setCellStyle(style2);
                            if (soharVisa.get(temp) != null) {
                                text.append("<td border = '1 px solid black'>");
                                text.append((soharVisa.get(temp)));
                                text.append("</td>");
                                dataCell.setCellValue(soharVisa.get(temp));
                            } else {
                                dataCell.setCellValue(0);
                                text.append("<td border = '1 px solid black'>");
                                text.append("0");
                                text.append("</td>");
                            }
                            dataCell = dataRow.createCell(3);
                            dataCell.setCellStyle(style2);
                            if (soharMastercard.get(temp) != null) {
                                text.append("<td border = '1 px solid black'>");
                                text.append((soharMastercard.get(temp)));
                                text.append("</td>");
                                dataCell.setCellValue(soharMastercard.get(temp));
                            } else {
                                dataCell.setCellValue(0);
                                text.append("<td border = '1 px solid black'>");
                                text.append("0");
                                text.append("</td>");
                            }
                            dataCell = dataRow.createCell(4);
                            dataCell.setCellStyle(style2);
                            if (temp.equals("00")) {
                                temp = "000";
                                if (soharHpsSohar.get(temp) != null) {
                                    text.append("<td border = '1 px solid black'>");
                                    text.append((soharHpsSohar.get(temp)));
                                    text.append("</td>");
                                    dataCell.setCellValue(soharHpsSohar.get(temp));
                                } else {
                                    dataCell.setCellValue(0);
                                    text.append("<td border = '1 px solid black'>");
                                    text.append("0");
                                    text.append("</td>");
                                }
                                temp = "00";
                            } else if (temp.equals("00 Dispute Server")) {
                                temp = "000 Dispute Server";
                                if (soharHpsSohar.get(temp) != null) {
                                    text.append("<td border = '1 px solid black'>");
                                    text.append((soharHpsSohar.get(temp)));
                                    text.append("</td>");
                                    dataCell.setCellValue(soharHpsSohar.get(temp));
                                } else {
                                    dataCell.setCellValue(0);
                                    text.append("<td border = '1 px solid black'>");
                                    text.append("0");
                                    text.append("</td>");
                                }
                                temp = "00 Dispute Server";
                            } else {
                                if (soharHpsSohar.get(temp) != null) {
                                    text.append("<td border = '1 px solid black'>");
                                    text.append((soharHpsSohar.get(temp)));
                                    text.append("</td>");
                                    dataCell.setCellValue(soharHpsSohar.get(temp));
                                } else {
                                    dataCell.setCellValue(0);
                                    text.append("<td border = '1 px solid black'>");
                                    text.append("0");
                                    text.append("</td>");
                                }
                            }
                            dataCell = dataRow.createCell(5);
                            dataCell.setCellStyle(style2);
                            if (soharVisionPlus.get(temp) != null) {
                                text.append("<td border = '1 px solid black'>");
                                text.append((soharVisionPlus.get(temp)));
                                text.append("</td>");
                                dataCell.setCellValue(soharVisionPlus.get(temp));
                            } else {
                                dataCell.setCellValue(0);
                                text.append("<td border = '1 px solid black'>");
                                text.append("0");
                                text.append("</td>");
                            }
                            dataCell = dataRow.createCell(6);
                            dataCell.setCellStyle(style2);
                            dataCell.setCellValue(soharCount.get(temp));
                            text.append("<td border = '1 px solid black'>");
                            text.append(soharCount.get(temp));
                            text.append("</td>");
                            text.append("</tr>");
 
                        }
 
                        if (tester == 0) {
                            rcText.append("<tr  border = '1 px solid black'>");
                            rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                            rcText.append("<b>-- Null --</b>" + "</td></tr>");
                        }
 
                        tester = 0;
                        rcText.append("</table>");
 
                        text.append("</table>");
                        for (int k = 0; k < 8; k++) {
                            sheet.autoSizeColumn(k);
                        }
                        tot = sheet.getLastRowNum();
                        AUBInitial = tot;
                        for (int i = soharInitial + 1; i < tot; i++) {
                            int j = i + 1;
                            RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                            RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                            RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                            RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                        }
                        for (int i = 0; i < 6; i++) {
                            int j = i + 1;
                            RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(soharInitial + 5, tot, i, j),
                                    sheet);
                            RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(soharInitial + 5, tot, i, j),
                                    sheet);
                            RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(soharInitial + 5, tot, i, j),
                                    sheet);
                            RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(soharInitial + 5, tot, i, j), sheet);
 
                        }
 
                        /*
                         * To maintain backup of report renaming previous report file
                         * with extention of timstamp
                         */
 
                        File oldInfileName = new File(prop.getProperty("soharInFile") + ".xlsx");
                        File newInfileName = new File(prop.getProperty("soharInFile") + timeStamp + ".xlsx");
 
                        oldInfileName.renameTo(newInfileName);
                        System.out.println("Rename completed successfully");
 
                        // *************** Code for comparing grandtotal for validating
                        // GG issue. ***********************
 
                        prop.setProperty("bshTotalCurr", Integer.toString(soharTotal[4]));
 
                        // ************** End of Code for comparing grandtotal for
                        // validating GG issue *******************
 
                    } else {
                        System.out.println("BSH Infile not present");
                    }
                } catch (Exception e) {
 
                    System.out.println("Exception in BSH report output...");
                    fin.close();
                    fo.close();
                    System.out.println(e);
 
                }
 
                // ********************************************************************************************************
 
                FileOutputStream fos = new FileOutputStream(new File(prop.getProperty("output") + timeStamp + ".xls"));
                book.write(fos);
                fos.flush();
                fos.close();
                System.out.println("written ");
 
                text.append("</tr>");
                text.append("</table>");
 
               // SendEmail email = new SendEmail();
             //   email.sendReportMail(text);
 
                // ************************ Pending Code
                // **********************************************
 
                /*
                 * // GG issue validating
                 *
                 * int result = ggIssueChecking();
                 *
                 * if(result==0) { email.sendGGIssue(ggIssue); }
                 */
 
                prop.setProperty("aboTotalPrev", prop.getProperty("aboTotalCurr"));
                prop.setProperty("afsTotalPrev", prop.getProperty("afsTotalCurr"));
                prop.setProperty("aubTotalPrev", prop.getProperty("aubTotalCurr"));
                prop.setProperty("bshTotalPrev", prop.getProperty("bshTotalCurr"));
 
                prop.store(fo, null);
                fo.close();
                fin.close();
       
    }
 
    public void transCountBahrain() throws IOException {
        // TODO Auto-generated method stub
        String timeStamp = new SimpleDateFormat("ddMM_HHmm").format(Calendar.getInstance().getTime());
        String timeStamp1 = new SimpleDateFormat("ddMM").format(Calendar.getInstance().getTime());
 
        System.out.println("Timestamp 1 :- " + timeStamp1);
 
        initialize();
 
        // Critical RC Null chcking flag
        int tester = 0;
 
        StringBuilder text = new StringBuilder();
 
        StringBuilder rcText = new StringBuilder();
 
        TreeMap<String, Integer> count = new TreeMap<>();
        TreeMap<String, Integer> visa = new TreeMap<>();
        TreeMap<String, Integer> mastercard = new TreeMap<>();
        TreeMap<String, Integer> omannet = new TreeMap<>();
        TreeMap<String, Integer> visionPlus = new TreeMap<>();
        TreeMap<String, Integer> other = new TreeMap<>();
        int tot;
        int bahrainInitial = 0, soharInitial = 0, AUBInitial = 0;
        int[] total = new int[5];
 
        /* Compare RC chart preparation */
 
        rcText.append("<table border='1' bordercolor='BLACK' style='border-collapse:collapse; font-family:Calibri;'>");
        rcText.append("<tr  border = '1 px solid black'>");
 
        rcText.append("<td colspan='1'  border = '1 px solid black'style ='text-align : center'>");
        rcText.append("<b>Response Code</b>" + "</td>");
        rcText.append("<td colspan='1'  border = '1 px solid black'style ='text-align : center'>");
        rcText.append("<b>Previous Count</b>" + "</td>");
        rcText.append("<td colspan='1'  border = '1 px solid black'style ='text-align : center'>");
        rcText.append("<b>Current Count</b>" + "</td>");
 
        rcText.append("</tr>");
 
        /* End of this section */
 
        HSSFWorkbook book = new HSSFWorkbook();
        HSSFSheet sheet = book.createSheet("transactions");
 
        HSSFCellStyle style = book.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont font = book.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        font.setFontName("Calibri");
        style.setFont(font);
        HSSFCellStyle style1 = book.createCellStyle();
        style1.setAlignment(HorizontalAlignment.LEFT);
        HSSFFont font1 = book.createFont();
        font1.setFontHeightInPoints((short) 12);
        font1.setBold(true);
        font1.setFontName("Calibri");
        style1.setFont(font1);
        HSSFCellStyle style2 = book.createCellStyle();
        HSSFFont font2 = book.createFont();
        font2.setFontHeightInPoints((short) 12);
        // font2.setBoldweight((short) 1000000);
        font2.setFontName("Calibri");
        style2.setAlignment(HorizontalAlignment.LEFT);
        style2.setFont(font2);
        // ******************************************** bahrain code
                // *************************************************************
                /*new coommand
                 * try {
                    if (new File(prop.getProperty("bahrainInFile") + ".xlsx").exists()) {
 
                        FileInputStream bahrainupdated = new FileInputStream(prop.getProperty("bahrainInFile") + ".xlsx");
 
                        XSSFWorkbook bahrainWorkbook = new XSSFWorkbook(bahrainupdated);
                        // FileOutputStream fos = new FileOutputStream(new
                        // File(prop.getProperty("bahrainBackup")));
                        // bahrainWorkbook.write(fos);
                        XSSFSheet bahrainSheeet = bahrainWorkbook.getSheet(bahrainWorkbook.getSheetName(0));
 
                        TreeMap<String, Integer> bahrainCount = new TreeMap<>();
                        TreeMap<String, Integer> bahrainVisa = new TreeMap<>();
                        TreeMap<String, Integer> bahrainMastercard = new TreeMap<>();
                        TreeMap<String, Integer> bahrainBenefit = new TreeMap<>();
                        TreeMap<String, Integer> bahrainAmex = new TreeMap<>();
                        TreeMap<String, Integer> bahrainJcb = new TreeMap<>();
                        TreeMap<String, Integer> bahrainOther = new TreeMap<>();
                        int[] bahrainTotal = new int[7];
 
                        rcText.append("<tr  border = '1 px solid black'>");
                        rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                        rcText.append("<b>========= AFS ===========</b>" + "</td></tr>");
 
                        for (int i = 8; i <= bahrainSheeet.getLastRowNum() - 2; i++) {
                            XSSFRow bahrainRow = bahrainSheeet.getRow(i);
                            // System.out.println(row.getCell(7));
                            XSSFCell cell = bahrainRow.getCell(18);
                            String code = cell.getStringCellValue();
                            XSSFCell cell2 = bahrainRow.getCell(17);
                            String interchange = cell2.getStringCellValue();
                            cell2 = bahrainRow.getCell(19);
                            String status = cell2.getStringCellValue();
                            if ((code.equals("") || code.equals(null) || code.equals(" ") || code.equals("0"))) {
                                if (bahrainRow.getCell(20).getStringCellValue().equalsIgnoreCase("In progress")) {
                                    code = new String("In progress");
                                } else {
                                    code = bahrainRow.getCell(20).getStringCellValue();
                                }
                            } else if (bahrainRow.getCell(20).getStringCellValue().equalsIgnoreCase("Timeout")) {
 
                                code = code + " " + new String(bahrainRow.getCell(20).getStringCellValue());
 
                            } else if (!(status.equals("Not initiated") || status.equals("Settled"))) {
                                code = code + " " + status;
                            }
                            switch (interchange) {
                            case "BENEFIT":
                                bahrainTotal[2]++;
                                if (bahrainBenefit.containsKey(code))
                                    bahrainBenefit.put(code, bahrainBenefit.get(code) + 1);
                                else
                                    bahrainBenefit.put(code, 1);
                                break;
                            case "MASTERCARD":
                                bahrainTotal[1]++;
                                if (bahrainMastercard.containsKey(code))
                                    bahrainMastercard.put(code, bahrainMastercard.get(code) + 1);
                                else
                                    bahrainMastercard.put(code, 1);
                                break;
                            case "VISA":
                                bahrainTotal[0]++;
                                if (bahrainVisa.containsKey(code))
                                    bahrainVisa.put(code, bahrainVisa.get(code) + 1);
                                else
                                    bahrainVisa.put(code, 1);
                                break;
                            case "AMEXACQUIRER":
                                bahrainTotal[3]++;
                                if (bahrainAmex.containsKey(code))
                                    bahrainAmex.put(code, bahrainAmex.get(code) + 1);
                                else
                                    bahrainAmex.put(code, 1);
                                break;
                            default:
                                if (bahrainOther.containsKey(code))
                                    bahrainOther.put(code, bahrainOther.get(code) + 1);
                                else
                                    bahrainOther.put(code, 1);
                            }
                            if (bahrainCount.containsKey(code)) {
                                bahrainCount.put(code, bahrainCount.get(code) + 1);
                            } else {
                                bahrainCount.put(code, 1);
                            }
                            bahrainTotal[4]++;
                        }
                        int bahrainSuccess = 0;
                        System.out.println(bahrainCount);
                        if (bahrainCount.get("000") != null && bahrainCount.get("00") != null) {
                            bahrainSuccess = bahrainCount.get("00") + bahrainCount.get("000");
                            bahrainCount.put("00", bahrainSuccess);
                            bahrainCount.remove("000");
                        } else if (bahrainCount.get("000") == null && bahrainCount.get("00") == null) {
 
                        } else {
                            if (bahrainCount.get("00") != null) {
                                bahrainSuccess = bahrainCount.get("00");
                            } else
                                bahrainSuccess = bahrainCount.get("000");
                            bahrainCount.put("00", bahrainSuccess);
                            bahrainCount.remove("000");
                        }
 
                        HSSFRow bahrainHeaderRow = sheet.createRow(sheet.getLastRowNum() + 2);
                        HSSFCell bahrainHeaderCell = bahrainHeaderRow.createCell(0);
                        bahrainHeaderCell.setCellValue("Transactions Detail report - Bahrain");
                        sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 6));
                        bahrainHeaderCell.setCellStyle(style);
                        text.append("<br>");
                        text.append("<br>");
                        text.append(
                                "<table border='1' bordercolor='BLACK' style='border-collapse:collapse; font-family:Calibri;'>");
                        text.append("<tr  border = '1 px solid black'>");
                        text.append("<td colspan='7'  border = '1 px solid black' style ='text-align : center'>");
                        text.append("<b>Transactions Detail report - Bahrain</b>");
                        text.append("</td>");
                        text.append("</tr>");
                        sheet.autoSizeColumn(0);
                        bahrainHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
 
                        // Makarand code for new output Sheet
                        String bankAFP = "Financial Institution ID : AFP - Arab Financial Serv";
                        String tranDateAFP = "Transaction Date : " + getCurrentBahrainTime(); // append
                                                                                                // Bahrain
                                                                                                // Time
                        String runDateAFP = "Run Date/Time : " + java.time.LocalDate.now() + " " + getCurrentBahrainTime()
                                + " Asia/Muscat"; // Append IST Date and Oman Time
 
                        for (int i = 1; i < 4; i++) {
                            bahrainHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
                            HSSFCell bankCell = bahrainHeaderRow.createCell(0);
                            sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 0, 6));
                            bankCell.setCellStyle(style1);
 
                            if (i == 1) {
 
                                bankCell.setCellValue(bankAFP);
                                text.append("<tr>");
                                text.append("<td colspan='7'   border = '1 px solid black'>");
                                text.append("<b>" + bankAFP + "</b>");
                                text.append("</td>");
                                text.append("</tr>");
                            }
 
                            if (i == 2) {
 
                                bankCell.setCellValue(tranDateAFP);
                                text.append("<tr>");
                                text.append("<td colspan='7'   border = '1 px solid black'>");
                                text.append("<b>" + tranDateAFP + "</b>");
                                text.append("</td>");
                                text.append("</tr>");
                            }
 
                            if (i == 3) {
 
                                bankCell.setCellValue(runDateAFP);
                                text.append("<tr>");
                                text.append("<td colspan='7'   border = '1 px solid black'>");
                                text.append("<b>" + runDateAFP + "</b>");
                                text.append("</td>");
                                text.append("</tr>");
                            }
 
                        }new command*/
 
                        /*
                         * for (int i = 2; i <= 5; i++) { text.append("<tr>");
                         * text.append("<td colspan='7'   border = '1 px solid black'>"
                         * ); bahrainHeaderRow = sheet.createRow(sheet.getLastRowNum() +
                         * 1); sheet.addMergedRegion(new
                         * CellRangeAddress(sheet.getLastRowNum(),
                         * sheet.getLastRowNum(), 0, 6)); HSSFCell bankCell =
                         * bahrainHeaderRow.createCell(0); String data; if (i == 5) {
                         * String tempData =
                         * bahrainSheeet.getRow(i).getCell(0).getStringCellValue(); data
                         * = tempData.substring(0, tempData.lastIndexOf('/')); data =
                         * data + "/Muscat";
                         *
                         * } else { data =
                         * bahrainSheeet.getRow(i).getCell(0).getStringCellValue(); }
                         * bankCell.setCellValue(data); // sheet.addMergedRegion(new
                         * CellRangeAddress(i, i, 0, 6)); bankCell.setCellStyle(style1);
                         * bankCell.setCellValue(data); text.append("<b>" + data +
                         * "</b>"); text.append("</td>"); text.append("</tr>"); }
                         */
 
                        /*new command
                        text.append("<tr>");
 
                        HSSFRow bahrainRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                        HSSFCell BahrainCellTotal = bahrainRowTotal.createCell(0);
                        BahrainCellTotal.setCellStyle(style);
                        sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 1));
                        BahrainCellTotal.setCellValue("Total Transactions");
                        text.append("<td border = '1 px solid black' colspan='2' rowspan='2'style ='text-align : center'>");
                        text.append("<b>Total Transactions</b>");
                        text.append("</td>");
                        BahrainCellTotal = bahrainRowTotal.createCell(2);
                        BahrainCellTotal.setCellStyle(style);
                        BahrainCellTotal.setCellValue("VISA");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>VISA</b>");
                        text.append("</td>");
                        BahrainCellTotal = bahrainRowTotal.createCell(3);
                        BahrainCellTotal.setCellStyle(style);
                        BahrainCellTotal.setCellValue("MASTER CARD");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>Master Card</b>");
                        text.append("</td>");
                        BahrainCellTotal = bahrainRowTotal.createCell(4);
                        BahrainCellTotal.setCellStyle(style);
                        BahrainCellTotal.setCellValue("BENEFIT");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>BENEFIT</b>");
                        text.append("</td>");
                        BahrainCellTotal = bahrainRowTotal.createCell(5);
                        BahrainCellTotal.setCellStyle(style);
                        BahrainCellTotal.setCellValue("AMEX");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>AMEX</b>");
                        text.append("</td>");
                        BahrainCellTotal = bahrainRowTotal.createCell(6);
                        BahrainCellTotal.setCellStyle(style);
                        BahrainCellTotal.setCellValue("GRAND TOTAL");
                        BahrainCellTotal.setCellStyle(style);
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>GRAND TOTAL</b>");
                        text.append("</td>");
                        text.append("</tr>");
                        text.append("<tr>");
                        bahrainRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                        for (int i = 0; i < bahrainTotal.length; i++) {
                            int j = i + 2;
                            BahrainCellTotal = bahrainRowTotal.createCell(j);
                            BahrainCellTotal.setCellStyle(style2);
                            BahrainCellTotal.setCellValue(bahrainTotal[i]);
                            text.append("<td  border = '1 px solid black'>");
                            text.append(bahrainTotal[i]);
                            text.append("</td>");
                        }
 
                        text.append("</tr>");
                        text.append("<tr>");
                        HSSFRow bahrainRowBlank = sheet.createRow(sheet.getLastRowNum() + 1);
                        HSSFCell baharinCellHeader = bahrainRowBlank.createCell(0);
                        baharinCellHeader.setCellStyle(style);
                        baharinCellHeader.setCellValue("Response code");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>Response code</b>");
                        text.append("</td>");
                        baharinCellHeader = bahrainRowBlank.createCell(1);
                        baharinCellHeader.setCellStyle(style);
                        baharinCellHeader.setCellValue("Response code description");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>Response code description</b>");
                        text.append("</td>");
                        baharinCellHeader = bahrainRowBlank.createCell(2);
                        baharinCellHeader.setCellStyle(style);
                        baharinCellHeader.setCellValue("VISA");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>VISA</b>");
                        text.append("</td>");
                        baharinCellHeader = bahrainRowBlank.createCell(3);
                        baharinCellHeader.setCellStyle(style);
                        baharinCellHeader.setCellValue("MASTERCARD");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>MASTERCARD</b>");
                        text.append("</td>");
                        baharinCellHeader = bahrainRowBlank.createCell(4);
                        baharinCellHeader.setCellStyle(style);
                        baharinCellHeader.setCellValue("BENEFIT");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>BENEFIT</b>");
                        text.append("</td>");
                        baharinCellHeader = bahrainRowBlank.createCell(5);
                        baharinCellHeader.setCellStyle(style);
                        baharinCellHeader.setCellValue("AMEX");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>AMEX</b>");
                        text.append("</td>");
                        baharinCellHeader = bahrainRowBlank.createCell(6);
                        baharinCellHeader.setCellStyle(style);
                        baharinCellHeader.setCellValue("Grand Total");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>Grand Total</b>");
                        text.append("</td>");
                        text.append("</tr>");
                        sheet.autoSizeColumn(0);
                        sheet.autoSizeColumn(1);
                        sheet.autoSizeColumn(2);
                        sheet.autoSizeColumn(3);
                        sheet.autoSizeColumn(4);
                        sheet.autoSizeColumn(5);
                        sheet.autoSizeColumn(6);
 
                        int bahrainLength = sheet.getLastRowNum() + 1;
                        Set<String> bahrainKeys = bahrainCount.keySet();
                        Iterator<String> bahrainIt = bahrainKeys.iterator();
                        String bahrainSettlementStatus = "";
                        for (int i = 2; i <= bahrainCount.size() + 1; i++) {
                            text.append("<tr>");
                            text.append("<td border = '1 px solid black'>");
                            String temp = bahrainIt.next();
                            HSSFRow dataRow = sheet.createRow(bahrainLength++);
                            dataRow.setRowStyle(style1);
                            HSSFCell dataCell = dataRow.createCell(0);
                            dataCell.setCellStyle(style2);
                            dataCell.setCellValue(temp);
                            text.append(temp);
                            text.append("</td>");
                            if (temp.length() > 4) {
                                bahrainSettlementStatus = temp.substring(3, temp.length());
                                System.out.println(bahrainSettlementStatus);
                            }
                            dataCell = dataRow.createCell(1);
                            dataCell.setCellStyle(style2);
                            if (prop.getProperty(temp) != null) {
                                dataCell.setCellValue(prop.getProperty(temp));
                                text.append("<td border = '1 px solid black'>");
                                text.append(prop.getProperty(temp));
                                text.append("</td>");
                            } else {
                                dataCell.setCellValue("unknown response code/ unknown status");
                                text.append("<td border = '1 px solid black'>");
                                text.append("unknown response code/ unknown status");
                                text.append("</td>");
                            }
                            dataCell = dataRow.createCell(2);
                            dataCell.setCellStyle(style2);
                            if (bahrainVisa.get(temp) != null) {
                                text.append("<td border = '1 px solid black'>");
                                text.append((bahrainVisa.get(temp)));
                                text.append("</td>");
                                dataCell.setCellValue(bahrainVisa.get(temp));
                            } else {
                                dataCell.setCellValue(0);
                                text.append("<td border = '1 px solid black'>");
                                text.append("0");
                                text.append("</td>");
                            }
                            dataCell = dataRow.createCell(3);
                            dataCell.setCellStyle(style2);
                            if (bahrainMastercard.get(temp) != null) {
                                text.append("<td border = '1 px solid black'>");
                                text.append((bahrainMastercard.get(temp)));
                                text.append("</td>");
                                dataCell.setCellValue(bahrainMastercard.get(temp));
                            } else {
                                dataCell.setCellValue(0);
                                text.append("<td border = '1 px solid black'>");
                                text.append("0");
                                text.append("</td>");
                            }
                            dataCell = dataRow.createCell(4);
                            dataCell.setCellStyle(style2);
                            if (bahrainBenefit.get(temp) != null) {
                                text.append("<td border = '1 px solid black'>");
                                text.append((bahrainBenefit.get(temp)));
                                text.append("</td>");
                                dataCell.setCellValue(bahrainBenefit.get(temp));
                            } else {
                                dataCell.setCellValue(0);
                                text.append("<td border = '1 px solid black'>");
                                text.append("0");
                                text.append("</td>");
                            }
                            dataCell = dataRow.createCell(5);
                            dataCell.setCellStyle(style2);
                            if (temp.equals("00")) {
                                temp = "000";
                                if (bahrainAmex.get(temp) != null) {
                                    text.append("<td border = '1 px solid black'>");
                                    text.append((bahrainAmex.get(temp)));
                                    text.append("</td>");
                                    dataCell.setCellValue(bahrainAmex.get(temp));
                                } else {
                                    dataCell.setCellValue(0);
                                    text.append("<td border = '1 px solid black'>");
                                    text.append("0");
                                    text.append("</td>");
                                }
                                temp = "00";
                            } else if (temp.equals("00 Dispute Server")) {
                                if (bahrainAmex.get(temp) != null) {
                                    text.append("<td border = '1 px solid black'>");
                                    text.append((bahrainAmex.get(temp)));
                                    text.append("</td>");
                                    dataCell.setCellValue(bahrainAmex.get(temp));
                                } else {
                                    dataCell.setCellValue(0);
                                    text.append("<td border = '1 px solid black'>");
                                    text.append("0");
                                    text.append("</td>");
                                }
                                temp = "00 Dispute Server";
                            } else {
                                if (bahrainAmex.get(temp) != null) {
                                    text.append("<td border = '1 px solid black'>");
                                    text.append((bahrainAmex.get(temp)));
                                    text.append("</td>");
                                    dataCell.setCellValue(bahrainAmex.get(temp));
                                } else {
                                    dataCell.setCellValue(0);
                                    text.append("<td border = '1 px solid black'>");
                                    text.append("0");
                                    text.append("</td>");
                                }
                            }
                            dataCell = dataRow.createCell(6);
                            dataCell.setCellStyle(style2);
                            dataCell.setCellValue(bahrainCount.get(temp));
                            text.append("<td border = '1 px solid black'>");
                            text.append(bahrainCount.get(temp));
                            text.append("</td>");
                            text.append("</tr>");
 
                        }
 
                        if (tester == 0) {
                            rcText.append("<tr  border = '1 px solid black'>");
                            rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                            rcText.append("<b>-- Null --</b>" + "</td></tr>");
                        }
 
                        tester = 0;
 
                        text.append("</table>");
                        text.append("<br>");
                        for (int k = 0; k < 8; k++) {
                            sheet.autoSizeColumn(k);
                        }
                        tot = sheet.getLastRowNum();
                        soharInitial = tot + 1;
                        for (int i = bahrainInitial + 1; i < tot; i++) {
                            int j = i + 1;
                            RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                            RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                            RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                            RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 6), sheet);
                        }
                        for (int i = 0; i < 6; i++) {
                            int j = i + 1;
                            RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(bahrainInitial + 6, tot, i, j),
                                    sheet);
                            RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(bahrainInitial + 6, tot, i, j),
                                    sheet);
                            RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(bahrainInitial + 6, tot, i, j),
                                    sheet);
                            RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(bahrainInitial + 6, tot, i, j),
                                    sheet);
 
                        }
 
                        // To maintain backup of report renaming previous report file
                        // with extention of timstamp
 
                        File oldInfileName = new File(prop.getProperty("bahrainInFile") + ".xlsx");
                        File newInfileName = new File(prop.getProperty("bahrainInFile") + timeStamp + ".xlsx");
 
                        oldInfileName.renameTo(newInfileName);
                        System.out.println("Rename completed successfully");
 
                    }
                } catch (Exception e) {
 
                    System.out.println("Exception in AFS report output...");
                    fin.close();
                    fo.close();
                    System.out.println(e);
                }*/
 
                // ********************************************* AUB CODE
                // *************************************************************
       
        
        try {
            if (new File(prop.getProperty("bahrainInFile") + ".xlsx").exists()) {

                FileInputStream bahrainUpdated = new FileInputStream(prop.getProperty("bahrainInFile") + ".xlsx");
                XSSFWorkbook bahrainWorkbook = new XSSFWorkbook(bahrainUpdated);
                // FileOutputStream fos = new FileOutputStream(new
                // File(prop.getProperty("bahrainBackup")));
                // bahrainWorkbook.write(fos);
                XSSFSheet bahrainSheeet = bahrainWorkbook.getSheet(bahrainWorkbook.getSheetName(0));
                TreeMap<String, Integer> bahrainCount = new TreeMap<>();
                TreeMap<String, Integer> bahrainVisa = new TreeMap<>();
                TreeMap<String, Integer> bahrainMastercard = new TreeMap<>();
                TreeMap<String, Integer> bahrainBenefit = new TreeMap<>();
                TreeMap<String, Integer> bahrainAmex = new TreeMap<>();
                TreeMap<String, Integer> bahrainVisionplus = new TreeMap<>();
                TreeMap<String, Integer> bahrainJcb = new TreeMap<>();
                TreeMap<String, Integer> bahrainOther = new TreeMap<>();
                int[] bahrainTotal = new int[7];
                // System.out.println(sheeet.getLastRowNum());

                rcText.append("<tr  border = '1 px solid black'>");
                rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                rcText.append("<b>========= bahrain ===========</b>" + "</td></tr>");

                for (int i = 8; i <= bahrainSheeet.getLastRowNum() - 2; i++) {
                    XSSFRow row = bahrainSheeet.getRow(i);
                    XSSFCell cell = row.getCell(18);
                    String code = cell.getStringCellValue();
                    XSSFCell cell2 = row.getCell(17);
                    String interchange = cell2.getStringCellValue();
                    cell2 = row.getCell(19);
                    String status = cell2.getStringCellValue();
                    if ((code.equals("") || code.equals(null) || code.equals(" ") || code.equals("0"))) {
                        if (row.getCell(20).getStringCellValue().equalsIgnoreCase("In progress")) {
                            code = new String("In progress");
                        } else {
                            code = row.getCell(20).getStringCellValue();
                        }
                    } else if (row.getCell(20).getStringCellValue().equalsIgnoreCase("Timeout")) {

                        code = code + " " + new String(row.getCell(20).getStringCellValue());

                    } else if (!(status.equals("Not initiated") || status.equals("Settled"))) {
                        System.out.println("here" + status);
                        System.out.println("RRn    " + row.getCell(6).getStringCellValue());
                        code = code + " " + status;
                        System.out.println("code");
                    }

                    switch (interchange) {
                    case "BENEFIT":
                        bahrainTotal[2]++;
                        if (bahrainBenefit.containsKey(code))
                            bahrainBenefit.put(code, bahrainBenefit.get(code) + 1);
                        else
                            bahrainBenefit.put(code, 1);
                        break;
                    case "MASTER CARD":
                        bahrainTotal[1]++;
                        if (bahrainMastercard.containsKey(code))
                            bahrainMastercard.put(code, bahrainMastercard.get(code) + 1);
                        else
                            bahrainMastercard.put(code, 1);
                        break;
                    case "VISA":
                        bahrainTotal[0]++;
                        if (bahrainVisa.containsKey(code))
                            bahrainVisa.put(code, bahrainVisa.get(code) + 1);
                        else
                            bahrainVisa.put(code, 1);
                        break;
                    case "AMEX":
                        bahrainTotal[3]++;
                        if (bahrainAmex.containsKey(code))
                            bahrainAmex.put(code, bahrainAmex.get(code) + 1);
                        else
                            bahrainAmex.put(code, 1);
                        break;
                    case "VISIONPLUSHOST":
                        bahrainTotal[4]++;
                        if (bahrainVisionplus.containsKey(code))
                            bahrainVisionplus.put(code, bahrainVisionplus.get(code) + 1);
                        else
                            bahrainVisionplus.put(code, 1);
                        break;
                    case "JCB":
                        bahrainTotal[5]++;
                        if (bahrainJcb.containsKey(code))
                            bahrainJcb.put(code, bahrainJcb.get(code) + 1);
                        else
                            bahrainJcb.put(code, 1);
                        break;
                    default:
                        if (bahrainOther.containsKey(code))
                            bahrainOther.put(code, bahrainOther.get(code) + 1);
                        else
                            bahrainOther.put(code, 1);
                    }

                    if (bahrainCount.containsKey(code)) {
                        bahrainCount.put(code, bahrainCount.get(code) + 1);
                    } else {
                        bahrainCount.put(code, 1);
                    }
                    bahrainTotal[6]++;
                }
                System.out.println("bahrain Count: " + bahrainCount);

                int bahrainSuccess = 0;

                if (bahrainCount.get("000") != null && bahrainCount.get("00") != null) {
                    bahrainSuccess = bahrainCount.get("00") + bahrainCount.get("000");
                    bahrainCount.put("00", bahrainSuccess);
                    bahrainCount.remove("000");
                } else if (bahrainCount.get("000") == null && bahrainCount.get("00") == null) {

                } else {
                    if (bahrainCount.get("00") != null) {
                        bahrainSuccess = bahrainCount.get("00");
                    } else
                        bahrainSuccess = bahrainCount.get("000");
                    bahrainCount.put("00", bahrainSuccess);
                    bahrainCount.remove("000");
                }

                System.out.println(bahrainCount);
                HSSFRow bahrainHeaderRow = sheet.createRow(sheet.getLastRowNum() + 2);
                HSSFCell bahrainHeaderCell = bahrainHeaderRow.createCell(0);
                bahrainHeaderCell.setCellValue("Transactions Detail report - bahrain");
                sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 8));
                bahrainHeaderCell.setCellStyle(style);
                text.append("<br>");
                text.append(
                        "<table border='1'  border = '1 px solid black' bordercolor='BLACK' style='border-collapse:collapse; font-family:Calibri;'>");
                text.append("<tr  border = '1 px solid black'>");
                text.append("<td colspan='9'  border = '1 px solid black'style ='text-align : center'>");
                text.append("<b>Transactions Detail report - bahrain</b>");
                text.append("</td>");
                text.append("</tr>");
                sheet.autoSizeColumn(0);
                bahrainHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);

                // Makarand code for new output Sheet
                String bankbahrain = "Financial Institution ID : ALI - Ahli United Bank";
                String tranDatebahrain = "Transaction Date : " + getCurrentBahrainTime(); // append
                                                                                        // Bahrain
                                                                                        // Time
                String runDatebahrain = "Run Date/Time : " + java.time.LocalDate.now() + " " + getCurrentBahrainTime()
                        + " Asia/Muscat"; // Append IST Date and Oman Time

                for (int i = 1; i < 4; i++) {
                    bahrainHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    HSSFCell bankCell = bahrainHeaderRow.createCell(0);
                    sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 0, 8));
                    bankCell.setCellStyle(style1);

                    if (i == 1) {

                        bankCell.setCellValue(bankbahrain);
                        // sheet.addMergedRegion(new
                        // CellRangeAddress(sheet.getLastRowNum(),
                        // sheet.getLastRowNum(), 0, 8));
                        // bankCell.setCellStyle(style1);
                        text.append("<tr>");
                        text.append("<td colspan='9'   border = '1 px solid black'>");
                        text.append("<b>" + bankbahrain + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }

                    if (i == 2) {

                        bankCell.setCellValue(tranDatebahrain);
                        // sheet.addMergedRegion(new
                        // CellRangeAddress(sheet.getLastRowNum(),
                        // sheet.getLastRowNum(), 0, 8));
                        // bankCell.setCellStyle(style1);
                        text.append("<tr>");
                        text.append("<td colspan='9'   border = '1 px solid black'>");
                        text.append("<b>" + tranDatebahrain + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }

                    if (i == 3) {

                        bankCell.setCellValue(runDatebahrain);
                        // sheet.addMergedRegion(new
                        // CellRangeAddress(sheet.getLastRowNum(),
                        // sheet.getLastRowNum(), 0, 8));
                        // bankCell.setCellStyle(style1);
                        text.append("<tr>");
                        text.append("<td colspan='9'   border = '1 px solid black'>");
                        text.append("<b>" + runDatebahrain + "</b>");
                        text.append("</td>");
                        text.append("</tr>");
                    }

                }

                /*
                 * for (int i = 2; i <= 5; i++) { bahrainHeaderRow =
                 * sheet.createRow(sheet.getLastRowNum() + 1); HSSFCell bankCell
                 * = bahrainHeaderRow.createCell(0); String data; if (i == 5) {
                 * String tempData =
                 * bahrainSheeet.getRow(i).getCell(0).getStringCellValue(); //
                 * System.out.println(tempData.lastIndexOf('/'));
                 *
                 * data = tempData.substring(0, tempData.lastIndexOf('/')); //
                 * System.out.println(data.concat("/Muscat")); data = data +
                 * "/Muscat";
                 *
                 * } else { data =
                 * bahrainSheeet.getRow(i).getCell(0).getStringCellValue(); } //
                 * System.out.println(data + " " + i);
                 * bankCell.setCellValue(data); sheet.addMergedRegion(new
                 * CellRangeAddress(sheet.getLastRowNum(),
                 * sheet.getLastRowNum(), 0, 8)); bankCell.setCellStyle(style1);
                 * bankCell.setCellValue(data); text.append("<tr>");
                 * text.append("<td colspan='9'   border = '1 px solid black'>"
                 * ); text.append("<b>" + data + "</b>"); text.append("</td>");
                 * text.append("</tr>"); }
                 */

                HSSFRow bahrainRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                HSSFCell bahrainCellTotal = bahrainRowTotal.createCell(0);
                bahrainCellTotal.setCellStyle(style);
                sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 1));
                bahrainCellTotal.setCellValue("Total Transactions");
                text.append("<tr>");
                text.append("<td border = '1 px solid black' colspan='2' rowspan='2'style ='text-align : center'>");
                text.append("<b>Total Transactions</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(2);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("VISA");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISA</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(3);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("MASTER CARD");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Master Card</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(4);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("BENEFIT");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>BENEFIT</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(5);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("AMEX");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>AMEX</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(6);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("VISION PLUS");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISION PLUS</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(7);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("JCB");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>JCB</b>");
                text.append("</td>");
                bahrainCellTotal = bahrainRowTotal.createCell(8);
                bahrainCellTotal.setCellStyle(style);
                bahrainCellTotal.setCellValue("GRAND TOTAL");
                bahrainRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                text.append("<td border = '1 px solid black'>");
                text.append("<b>GRAND TOTAL</b>");
                text.append("</td>");
                text.append("</tr>");
                text.append("<tr>");
                for (int i = 0; i < bahrainTotal.length; i++) {
                    int j = i + 2;
                    bahrainCellTotal = bahrainRowTotal.createCell(j);
                    bahrainCellTotal.setCellStyle(style2);
                    bahrainCellTotal.setCellValue(bahrainTotal[i]);
                    text.append("<td  border = '1 px solid black'>");
                    text.append(bahrainTotal[i]);
                    text.append("</td>");
                }

                text.append("</tr>");
                text.append("<tr>");
                HSSFRow bahrainRowBlank = sheet.createRow(sheet.getLastRowNum() + 1);
                HSSFCell bahrainCellHeader = bahrainRowBlank.createCell(0);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("Response code");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Response code</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(1);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("Response code description");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Response code description/<b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(2);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("VISA");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISA</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(3);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("MASTERCARD");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>MASTERCARD</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(4);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("BENEFIT");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>BENEFIT</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(5);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("AMEX");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>AMEX</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(6);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("VISION PLUS");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>VISION PLUS</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(7);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("JCB");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>JCB</b>");
                text.append("</td>");
                bahrainCellHeader = bahrainRowBlank.createCell(8);
                bahrainCellHeader.setCellStyle(style);
                bahrainCellHeader.setCellValue("Grand Total");
                text.append("<td border = '1 px solid black'>");
                text.append("<b>Grand Total</b>");
                text.append("</td>");
                text.append("</tr>");

                sheet.autoSizeColumn(0);
                sheet.autoSizeColumn(1);
                sheet.autoSizeColumn(2);
                sheet.autoSizeColumn(3);
                sheet.autoSizeColumn(4);
                sheet.autoSizeColumn(5);
                sheet.autoSizeColumn(6);
                sheet.autoSizeColumn(7);
                sheet.autoSizeColumn(8);
                int bahrainLength = sheet.getLastRowNum() + 1;

                Set<String> bahrainKeys = bahrainCount.keySet();
                Iterator<String> bahrainIt = bahrainKeys.iterator();
                String settlementStatus = "";
                for (int i = 2; i <= bahrainCount.size() + 1; i++) {
                    text.append("<tr>");
                    String temp = bahrainIt.next();
                    HSSFRow dataRow = sheet.createRow(bahrainLength++);
                    dataRow.setRowStyle(style);
                    HSSFCell dataCell = dataRow.createCell(0);
                    dataCell.setCellStyle(style2);
                    dataCell.setCellValue(temp);
                    text.append("<td border = '1 px solid black'>");
                    text.append(temp);
                    text.append("</td>");

                    if (temp.length() > 4) {
                        settlementStatus = temp.substring(3, temp.length());
                        System.out.println(settlementStatus);
                    }

                    dataCell = dataRow.createCell(1);
                    dataCell.setCellStyle(style2);

                    if (prop.getProperty(temp) != null) {
                        dataCell.setCellValue(prop.getProperty(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append(prop.getProperty(temp));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue("unknown response code/ unknown status");
                        text.append("<td border = '1 px solid black'>");
                        text.append("unknown response code/ unknown status");
                        text.append("</td>");
                    }
                    // }
                    dataCell = dataRow.createCell(2);
                    dataCell.setCellStyle(style2);
                    if (bahrainVisa.get(temp) != null) {
                        dataCell.setCellValue(bahrainVisa.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((bahrainVisa.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(3);
                    dataCell.setCellStyle(style2);
                    if (bahrainMastercard.get(temp) != null) {
                        dataCell.setCellValue(bahrainMastercard.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((bahrainMastercard.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(4);
                    dataCell.setCellStyle(style2);
                    if (bahrainBenefit.get(temp) != null) {
                        dataCell.setCellValue(bahrainBenefit.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((bahrainBenefit.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(5);
                    dataCell.setCellStyle(style2);
                    if (temp.equals("00")) {
                        temp = "000";
                        if (bahrainAmex.get(temp) != null) {
                            dataCell.setCellValue(bahrainAmex.get(temp));
                            text.append("<td border = '1 px solid black'>");
                            text.append((bahrainAmex.get(temp)));
                            text.append("</td>");
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00";
                    } else if (temp.equals("00 Dispute Server")) {
                        if (bahrainAmex.get(temp) != null) {
                            dataCell.setCellValue(bahrainAmex.get(temp));
                            text.append("<td border = '1 px solid black'>");
                            text.append((bahrainAmex.get(temp)));
                            text.append("</td>");
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                        temp = "00 Dispute Server";
                    } else {
                        if (bahrainAmex.get(temp) != null) {
                            dataCell.setCellValue(bahrainAmex.get(temp));
                            text.append("<td border = '1 px solid black'>");
                            text.append((bahrainAmex.get(temp)));
                            text.append("</td>");
                        } else {
                            dataCell.setCellValue(0);
                            text.append("<td border = '1 px solid black'>");
                            text.append("0");
                            text.append("</td>");
                        }
                    }
                    dataCell = dataRow.createCell(6);
                    dataCell.setCellStyle(style2);
                    if (bahrainVisionplus.get(temp) != null) {
                        dataCell.setCellValue(bahrainVisionplus.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((bahrainVisionplus.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(7);
                    dataCell.setCellStyle(style2);
                    if (bahrainJcb.get(temp) != null) {
                        dataCell.setCellValue(bahrainJcb.get(temp));
                        text.append("<td border = '1 px solid black'>");
                        text.append((bahrainJcb.get(temp)));
                        text.append("</td>");
                    } else {
                        dataCell.setCellValue(0);
                        text.append("<td border = '1 px solid black'>");
                        text.append("0");
                        text.append("</td>");
                    }
                    dataCell = dataRow.createCell(8);
                    dataCell.setCellStyle(style2);
                    dataCell.setCellValue(bahrainCount.get(temp));
                    text.append("<td border = '1 px solid black'>");
                    text.append(bahrainCount.get(temp));
                    text.append("</td>");
                    text.append("</tr>");

                }

                /*
                 * if(tester == 0) { rcText.append(
                 * "<tr  border = '1 px solid black'>"); rcText.append(
                 * "<td colspan='3'  border = '1 px solid black'style ='text-align : center'>"
                 * ); rcText.append("<b>-- Null --</b>"+ "</td></tr>"); }
                 */
                tester = 0;

                text.append("</table>");
                text.append("<br>");
                tot = sheet.getLastRowNum();
                for (int i = bahrainInitial + 2; i < tot; i++) {
                    int j = i + 1;
                    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                }
                for (int i = 0; i < 8; i++) {
                    int j = i + 1;
                    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(bahrainInitial + 7, tot, i, j),
                            sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(bahrainInitial + 7, tot, i, j), sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(bahrainInitial + 7, tot, i, j), sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(bahrainInitial + 7, tot, i, j), sheet);

                }

                /*
                 * To maintain backup of report renaming previous report file
                 * with extention of timstamp
                 */

                File oldInfileName = new File(prop.getProperty("bahrainInFile") + ".xlsx");
                File newInfileName = new File(prop.getProperty("bahrainInFile") + timeStamp + ".xlsx");

                oldInfileName.renameTo(newInfileName);
                System.out.println("Rename completed successfully");

                // *************** Code for comparing grandtotal for validating
                // GG issue. ***********************

                // prop.setProperty("bahrainTotalCurr",Integer.toString(bahrainTotal[6]));

                // ************** End of Code for comparing grandtotal for
                // validating GG issue *******************

            } else {
                System.out.println("bahrain Infile not present");
            }
        } catch (Exception e) {

            System.out.println("Exception in bahrain report output...");
            fin.close();
            fo.close();
            System.out.println(e);
        }
        
        
       
        try {
                    if (new File(prop.getProperty("aubInFile") + ".xlsx").exists()) {
 
                        FileInputStream AUBUpdated = new FileInputStream(prop.getProperty("aubInFile") + ".xlsx");
                        XSSFWorkbook AUBWorkbook = new XSSFWorkbook(AUBUpdated);
                        // FileOutputStream fos = new FileOutputStream(new
                        // File(prop.getProperty("aubBackup")));
                        // AUBWorkbook.write(fos);
                        XSSFSheet AUBSheeet = AUBWorkbook.getSheet(AUBWorkbook.getSheetName(0));
                        TreeMap<String, Integer> AUBCount = new TreeMap<>();
                        TreeMap<String, Integer> AUBVisa = new TreeMap<>();
                        TreeMap<String, Integer> AUBMastercard = new TreeMap<>();
                        TreeMap<String, Integer> AUBBenefit = new TreeMap<>();
                        TreeMap<String, Integer> AUBAmex = new TreeMap<>();
                        TreeMap<String, Integer> AUBVisionplus = new TreeMap<>();
                        TreeMap<String, Integer> AUBJcb = new TreeMap<>();
                        TreeMap<String, Integer> AUBOther = new TreeMap<>();
                        int[] AUBTotal = new int[7];
                        // System.out.println(sheeet.getLastRowNum());
 
                        rcText.append("<tr  border = '1 px solid black'>");
                        rcText.append("<td colspan='3'  border = '1 px solid black'style ='text-align : center'>");
                        rcText.append("<b>========= AUB ===========</b>" + "</td></tr>");
 
                        for (int i = 8; i <= AUBSheeet.getLastRowNum() - 2; i++) {
                            XSSFRow row = AUBSheeet.getRow(i);
                            XSSFCell cell = row.getCell(18);
                            String code = cell.getStringCellValue();
                            XSSFCell cell2 = row.getCell(17);
                            String interchange = cell2.getStringCellValue();
                            cell2 = row.getCell(19);
                            String status = cell2.getStringCellValue();
                            if ((code.equals("") || code.equals(null) || code.equals(" ") || code.equals("0"))) {
                                if (row.getCell(20).getStringCellValue().equalsIgnoreCase("In progress")) {
                                    code = new String("In progress");
                                } else {
                                    code = row.getCell(20).getStringCellValue();
                                }
                            } else if (row.getCell(20).getStringCellValue().equalsIgnoreCase("Timeout")) {
 
                                code = code + " " + new String(row.getCell(20).getStringCellValue());
 
                            } else if (!(status.equals("Not initiated") || status.equals("Settled"))) {
                                System.out.println("here" + status);
                                System.out.println("RRn    " + row.getCell(6).getStringCellValue());
                                code = code + " " + status;
                                System.out.println("code");
                            }
 
                            switch (interchange) {
                            case "BENEFIT":
                                AUBTotal[2]++;
                                if (AUBBenefit.containsKey(code))
                                    AUBBenefit.put(code, AUBBenefit.get(code) + 1);
                                else
                                    AUBBenefit.put(code, 1);
                                break;
                            case "MASTER CARD":
                                AUBTotal[1]++;
                                if (AUBMastercard.containsKey(code))
                                    AUBMastercard.put(code, AUBMastercard.get(code) + 1);
                                else
                                    AUBMastercard.put(code, 1);
                                break;
                            case "VISA":
                                AUBTotal[0]++;
                                if (AUBVisa.containsKey(code))
                                    AUBVisa.put(code, AUBVisa.get(code) + 1);
                                else
                                    AUBVisa.put(code, 1);
                                break;
                            case "AMEX":
                                AUBTotal[3]++;
                                if (AUBAmex.containsKey(code))
                                    AUBAmex.put(code, AUBAmex.get(code) + 1);
                                else
                                    AUBAmex.put(code, 1);
                                break;
                            case "VISIONPLUSHOST":
                                AUBTotal[4]++;
                                if (AUBVisionplus.containsKey(code))
                                    AUBVisionplus.put(code, AUBVisionplus.get(code) + 1);
                                else
                                    AUBVisionplus.put(code, 1);
                                break;
                            case "JCB":
                                AUBTotal[5]++;
                                if (AUBJcb.containsKey(code))
                                    AUBJcb.put(code, AUBJcb.get(code) + 1);
                                else
                                    AUBJcb.put(code, 1);
                                break;
                            default:
                                if (AUBOther.containsKey(code))
                                    AUBOther.put(code, AUBOther.get(code) + 1);
                                else
                                    AUBOther.put(code, 1);
                            }
 
                            if (AUBCount.containsKey(code)) {
                                AUBCount.put(code, AUBCount.get(code) + 1);
                            } else {
                                AUBCount.put(code, 1);
                            }
                            AUBTotal[6]++;
                        }
                        System.out.println("AUB Count: " + AUBCount);
 
                        int AUBSuccess = 0;
 
                        if (AUBCount.get("000") != null && AUBCount.get("00") != null) {
                            AUBSuccess = AUBCount.get("00") + AUBCount.get("000");
                            AUBCount.put("00", AUBSuccess);
                            AUBCount.remove("000");
                        } else if (AUBCount.get("000") == null && AUBCount.get("00") == null) {
 
                        } else {
                            if (AUBCount.get("00") != null) {
                                AUBSuccess = AUBCount.get("00");
                            } else
                                AUBSuccess = AUBCount.get("000");
                            AUBCount.put("00", AUBSuccess);
                            AUBCount.remove("000");
                        }
 
                        System.out.println(AUBCount);
                        HSSFRow AUBHeaderRow = sheet.createRow(sheet.getLastRowNum() + 2);
                        HSSFCell AUBHeaderCell = AUBHeaderRow.createCell(0);
                        AUBHeaderCell.setCellValue("Transactions Detail report - AUB");
                        sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 8));
                        AUBHeaderCell.setCellStyle(style);
                        text.append("<br>");
                        text.append(
                                "<table border='1'  border = '1 px solid black' bordercolor='BLACK' style='border-collapse:collapse; font-family:Calibri;'>");
                        text.append("<tr  border = '1 px solid black'>");
                        text.append("<td colspan='9'  border = '1 px solid black'style ='text-align : center'>");
                        text.append("<b>Transactions Detail report - AUB</b>");
                        text.append("</td>");
                        text.append("</tr>");
                        sheet.autoSizeColumn(0);
                        AUBHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
 
                        // Makarand code for new output Sheet
                        String bankAUB = "Financial Institution ID : ALI - Ahli United Bank";
                        String tranDateAUB = "Transaction Date : " + getCurrentBahrainTime(); // append
                                                                                                // Bahrain
                                                                                                // Time
                        String runDateAUB = "Run Date/Time : " + java.time.LocalDate.now() + " " + getCurrentBahrainTime()
                                + " Asia/Muscat"; // Append IST Date and Oman Time
 
                        for (int i = 1; i < 4; i++) {
                            AUBHeaderRow = sheet.createRow(sheet.getLastRowNum() + 1);
                            HSSFCell bankCell = AUBHeaderRow.createCell(0);
                            sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum(), 0, 8));
                            bankCell.setCellStyle(style1);
 
                            if (i == 1) {
 
                                bankCell.setCellValue(bankAUB);
                                // sheet.addMergedRegion(new
                                // CellRangeAddress(sheet.getLastRowNum(),
                                // sheet.getLastRowNum(), 0, 8));
                                // bankCell.setCellStyle(style1);
                                text.append("<tr>");
                                text.append("<td colspan='9'   border = '1 px solid black'>");
                                text.append("<b>" + bankAUB + "</b>");
                                text.append("</td>");
                                text.append("</tr>");
                            }
 
                            if (i == 2) {
 
                                bankCell.setCellValue(tranDateAUB);
                                // sheet.addMergedRegion(new
                                // CellRangeAddress(sheet.getLastRowNum(),
                                // sheet.getLastRowNum(), 0, 8));
                                // bankCell.setCellStyle(style1);
                                text.append("<tr>");
                                text.append("<td colspan='9'   border = '1 px solid black'>");
                                text.append("<b>" + tranDateAUB + "</b>");
                                text.append("</td>");
                                text.append("</tr>");
                            }
 
                            if (i == 3) {
 
                                bankCell.setCellValue(runDateAUB);
                                // sheet.addMergedRegion(new
                                // CellRangeAddress(sheet.getLastRowNum(),
                                // sheet.getLastRowNum(), 0, 8));
                                // bankCell.setCellStyle(style1);
                                text.append("<tr>");
                                text.append("<td colspan='9'   border = '1 px solid black'>");
                                text.append("<b>" + runDateAUB + "</b>");
                                text.append("</td>");
                                text.append("</tr>");
                            }
 
                        }
 
                        /*
                         * for (int i = 2; i <= 5; i++) { AUBHeaderRow =
                         * sheet.createRow(sheet.getLastRowNum() + 1); HSSFCell bankCell
                         * = AUBHeaderRow.createCell(0); String data; if (i == 5) {
                         * String tempData =
                         * AUBSheeet.getRow(i).getCell(0).getStringCellValue(); //
                         * System.out.println(tempData.lastIndexOf('/'));
                         *
                         * data = tempData.substring(0, tempData.lastIndexOf('/')); //
                         * System.out.println(data.concat("/Muscat")); data = data +
                         * "/Muscat";
                         *
                         * } else { data =
                         * AUBSheeet.getRow(i).getCell(0).getStringCellValue(); } //
                         * System.out.println(data + " " + i);
                         * bankCell.setCellValue(data); sheet.addMergedRegion(new
                         * CellRangeAddress(sheet.getLastRowNum(),
                         * sheet.getLastRowNum(), 0, 8)); bankCell.setCellStyle(style1);
                         * bankCell.setCellValue(data); text.append("<tr>");
                         * text.append("<td colspan='9'   border = '1 px solid black'>"
                         * ); text.append("<b>" + data + "</b>"); text.append("</td>");
                         * text.append("</tr>"); }
                         */
 
                        HSSFRow AUBRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                        HSSFCell AUBCellTotal = AUBRowTotal.createCell(0);
                        AUBCellTotal.setCellStyle(style);
                        sheet.addMergedRegion(new CellRangeAddress(sheet.getLastRowNum(), sheet.getLastRowNum() + 1, 0, 1));
                        AUBCellTotal.setCellValue("Total Transactions");
                        text.append("<tr>");
                        text.append("<td border = '1 px solid black' colspan='2' rowspan='2'style ='text-align : center'>");
                        text.append("<b>Total Transactions</b>");
                        text.append("</td>");
                        AUBCellTotal = AUBRowTotal.createCell(2);
                        AUBCellTotal.setCellStyle(style);
                        AUBCellTotal.setCellValue("VISA");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>VISA</b>");
                        text.append("</td>");
                        AUBCellTotal = AUBRowTotal.createCell(3);
                        AUBCellTotal.setCellStyle(style);
                        AUBCellTotal.setCellValue("MASTER CARD");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>Master Card</b>");
                        text.append("</td>");
                        AUBCellTotal = AUBRowTotal.createCell(4);
                        AUBCellTotal.setCellStyle(style);
                        AUBCellTotal.setCellValue("BENEFIT");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>BENEFIT</b>");
                        text.append("</td>");
                        AUBCellTotal = AUBRowTotal.createCell(5);
                        AUBCellTotal.setCellStyle(style);
                        AUBCellTotal.setCellValue("AMEX");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>AMEX</b>");
                        text.append("</td>");
                        AUBCellTotal = AUBRowTotal.createCell(6);
                        AUBCellTotal.setCellStyle(style);
                        AUBCellTotal.setCellValue("VISION PLUS");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>VISION PLUS</b>");
                        text.append("</td>");
                        AUBCellTotal = AUBRowTotal.createCell(7);
                        AUBCellTotal.setCellStyle(style);
                        AUBCellTotal.setCellValue("JCB");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>JCB</b>");
                        text.append("</td>");
                        AUBCellTotal = AUBRowTotal.createCell(8);
                        AUBCellTotal.setCellStyle(style);
                        AUBCellTotal.setCellValue("GRAND TOTAL");
                        AUBRowTotal = sheet.createRow(sheet.getLastRowNum() + 1);
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>GRAND TOTAL</b>");
                        text.append("</td>");
                        text.append("</tr>");
                        text.append("<tr>");
                        for (int i = 0; i < AUBTotal.length; i++) {
                            int j = i + 2;
                            AUBCellTotal = AUBRowTotal.createCell(j);
                            AUBCellTotal.setCellStyle(style2);
                            AUBCellTotal.setCellValue(AUBTotal[i]);
                            text.append("<td  border = '1 px solid black'>");
                            text.append(AUBTotal[i]);
                            text.append("</td>");
                        }
 
                        text.append("</tr>");
                        text.append("<tr>");
                        HSSFRow AUBRowBlank = sheet.createRow(sheet.getLastRowNum() + 1);
                        HSSFCell AUBCellHeader = AUBRowBlank.createCell(0);
                        AUBCellHeader.setCellStyle(style);
                        AUBCellHeader.setCellValue("Response code");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>Response code</b>");
                        text.append("</td>");
                        AUBCellHeader = AUBRowBlank.createCell(1);
                        AUBCellHeader.setCellStyle(style);
                        AUBCellHeader.setCellValue("Response code description");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>Response code description/<b>");
                        text.append("</td>");
                        AUBCellHeader = AUBRowBlank.createCell(2);
                        AUBCellHeader.setCellStyle(style);
                        AUBCellHeader.setCellValue("VISA");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>VISA</b>");
                        text.append("</td>");
                        AUBCellHeader = AUBRowBlank.createCell(3);
                        AUBCellHeader.setCellStyle(style);
                        AUBCellHeader.setCellValue("MASTERCARD");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>MASTERCARD</b>");
                        text.append("</td>");
                        AUBCellHeader = AUBRowBlank.createCell(4);
                        AUBCellHeader.setCellStyle(style);
                        AUBCellHeader.setCellValue("BENEFIT");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>BENEFIT</b>");
                        text.append("</td>");
                        AUBCellHeader = AUBRowBlank.createCell(5);
                        AUBCellHeader.setCellStyle(style);
                        AUBCellHeader.setCellValue("AMEX");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>AMEX</b>");
                        text.append("</td>");
                        AUBCellHeader = AUBRowBlank.createCell(6);
                        AUBCellHeader.setCellStyle(style);
                        AUBCellHeader.setCellValue("VISION PLUS");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>VISION PLUS</b>");
                        text.append("</td>");
                        AUBCellHeader = AUBRowBlank.createCell(7);
                        AUBCellHeader.setCellStyle(style);
                        AUBCellHeader.setCellValue("JCB");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>JCB</b>");
                        text.append("</td>");
                        AUBCellHeader = AUBRowBlank.createCell(8);
                        AUBCellHeader.setCellStyle(style);
                        AUBCellHeader.setCellValue("Grand Total");
                        text.append("<td border = '1 px solid black'>");
                        text.append("<b>Grand Total</b>");
                        text.append("</td>");
                        text.append("</tr>");
 
                        sheet.autoSizeColumn(0);
                        sheet.autoSizeColumn(1);
                        sheet.autoSizeColumn(2);
                        sheet.autoSizeColumn(3);
                        sheet.autoSizeColumn(4);
                        sheet.autoSizeColumn(5);
                        sheet.autoSizeColumn(6);
                        sheet.autoSizeColumn(7);
                        sheet.autoSizeColumn(8);
                        int AUBLength = sheet.getLastRowNum() + 1;
 
                        Set<String> AUBKeys = AUBCount.keySet();
                        Iterator<String> AUBIt = AUBKeys.iterator();
                        String settlementStatus = "";
                        for (int i = 2; i <= AUBCount.size() + 1; i++) {
                            text.append("<tr>");
                            String temp = AUBIt.next();
                            HSSFRow dataRow = sheet.createRow(AUBLength++);
                            dataRow.setRowStyle(style);
                            HSSFCell dataCell = dataRow.createCell(0);
                            dataCell.setCellStyle(style2);
                            dataCell.setCellValue(temp);
                            text.append("<td border = '1 px solid black'>");
                            text.append(temp);
                            text.append("</td>");
 
                            if (temp.length() > 4) {
                                settlementStatus = temp.substring(3, temp.length());
                                System.out.println(settlementStatus);
                            }
 
                            dataCell = dataRow.createCell(1);
                            dataCell.setCellStyle(style2);
 
                            if (prop.getProperty(temp) != null) {
                                dataCell.setCellValue(prop.getProperty(temp));
                                text.append("<td border = '1 px solid black'>");
                                text.append(prop.getProperty(temp));
                                text.append("</td>");
                            } else {
                                dataCell.setCellValue("unknown response code/ unknown status");
                                text.append("<td border = '1 px solid black'>");
                                text.append("unknown response code/ unknown status");
                                text.append("</td>");
                            }
                            // }
                            dataCell = dataRow.createCell(2);
                            dataCell.setCellStyle(style2);
                            if (AUBVisa.get(temp) != null) {
                                dataCell.setCellValue(AUBVisa.get(temp));
                                text.append("<td border = '1 px solid black'>");
                                text.append((AUBVisa.get(temp)));
                                text.append("</td>");
                            } else {
                                dataCell.setCellValue(0);
                                text.append("<td border = '1 px solid black'>");
                                text.append("0");
                                text.append("</td>");
                            }
                            dataCell = dataRow.createCell(3);
                            dataCell.setCellStyle(style2);
                            if (AUBMastercard.get(temp) != null) {
                                dataCell.setCellValue(AUBMastercard.get(temp));
                                text.append("<td border = '1 px solid black'>");
                                text.append((AUBMastercard.get(temp)));
                                text.append("</td>");
                            } else {
                                dataCell.setCellValue(0);
                                text.append("<td border = '1 px solid black'>");
                                text.append("0");
                                text.append("</td>");
                            }
                            dataCell = dataRow.createCell(4);
                            dataCell.setCellStyle(style2);
                            if (AUBBenefit.get(temp) != null) {
                                dataCell.setCellValue(AUBBenefit.get(temp));
                                text.append("<td border = '1 px solid black'>");
                                text.append((AUBBenefit.get(temp)));
                                text.append("</td>");
                            } else {
                                dataCell.setCellValue(0);
                                text.append("<td border = '1 px solid black'>");
                                text.append("0");
                                text.append("</td>");
                            }
                            dataCell = dataRow.createCell(5);
                            dataCell.setCellStyle(style2);
                            if (temp.equals("00")) {
                                temp = "000";
                                if (AUBAmex.get(temp) != null) {
                                    dataCell.setCellValue(AUBAmex.get(temp));
                                    text.append("<td border = '1 px solid black'>");
                                    text.append((AUBAmex.get(temp)));
                                    text.append("</td>");
                                } else {
                                    dataCell.setCellValue(0);
                                    text.append("<td border = '1 px solid black'>");
                                    text.append("0");
                                    text.append("</td>");
                                }
                                temp = "00";
                            } else if (temp.equals("00 Dispute Server")) {
                                if (AUBAmex.get(temp) != null) {
                                    dataCell.setCellValue(AUBAmex.get(temp));
                                    text.append("<td border = '1 px solid black'>");
                                    text.append((AUBAmex.get(temp)));
                                    text.append("</td>");
                                } else {
                                    dataCell.setCellValue(0);
                                    text.append("<td border = '1 px solid black'>");
                                    text.append("0");
                                    text.append("</td>");
                                }
                                temp = "00 Dispute Server";
                            } else {
                                if (AUBAmex.get(temp) != null) {
                                    dataCell.setCellValue(AUBAmex.get(temp));
                                    text.append("<td border = '1 px solid black'>");
                                    text.append((AUBAmex.get(temp)));
                                    text.append("</td>");
                                } else {
                                    dataCell.setCellValue(0);
                                    text.append("<td border = '1 px solid black'>");
                                    text.append("0");
                                    text.append("</td>");
                                }
                            }
                            dataCell = dataRow.createCell(6);
                            dataCell.setCellStyle(style2);
                            if (AUBVisionplus.get(temp) != null) {
                                dataCell.setCellValue(AUBVisionplus.get(temp));
                                text.append("<td border = '1 px solid black'>");
                                text.append((AUBVisionplus.get(temp)));
                                text.append("</td>");
                            } else {
                                dataCell.setCellValue(0);
                                text.append("<td border = '1 px solid black'>");
                                text.append("0");
                                text.append("</td>");
                            }
                            dataCell = dataRow.createCell(7);
                            dataCell.setCellStyle(style2);
                            if (AUBJcb.get(temp) != null) {
                                dataCell.setCellValue(AUBJcb.get(temp));
                                text.append("<td border = '1 px solid black'>");
                                text.append((AUBJcb.get(temp)));
                                text.append("</td>");
                            } else {
                                dataCell.setCellValue(0);
                                text.append("<td border = '1 px solid black'>");
                                text.append("0");
                                text.append("</td>");
                            }
                            dataCell = dataRow.createCell(8);
                            dataCell.setCellStyle(style2);
                            dataCell.setCellValue(AUBCount.get(temp));
                            text.append("<td border = '1 px solid black'>");
                            text.append(AUBCount.get(temp));
                            text.append("</td>");
                            text.append("</tr>");
 
                        }
 
                        /*
                         * if(tester == 0) { rcText.append(
                         * "<tr  border = '1 px solid black'>"); rcText.append(
                         * "<td colspan='3'  border = '1 px solid black'style ='text-align : center'>"
                         * ); rcText.append("<b>-- Null --</b>"+ "</td></tr>"); }
                         */
                        tester = 0;
 
                        text.append("</table>");
                        text.append("<br>");
                        tot = sheet.getLastRowNum();
                        for (int i = AUBInitial + 2; i < tot; i++) {
                            int j = i + 1;
                            RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                            RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                            RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                            RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(i, j, 0, 8), sheet);
                        }
                        for (int i = 0; i < 8; i++) {
                            int j = i + 1;
                            RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(AUBInitial + 7, tot, i, j),
                                    sheet);
                            RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(AUBInitial + 7, tot, i, j), sheet);
                            RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(AUBInitial + 7, tot, i, j), sheet);
                            RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(AUBInitial + 7, tot, i, j), sheet);
 
                        }
 
                        /*
                         * To maintain backup of report renaming previous report file
                         * with extention of timstamp
                         */
 
                        File oldInfileName = new File(prop.getProperty("aubInFile") + ".xlsx");
                        File newInfileName = new File(prop.getProperty("aubInFile") + timeStamp + ".xlsx");
 
                        oldInfileName.renameTo(newInfileName);
                        System.out.println("Rename completed successfully");
 
                        // *************** Code for comparing grandtotal for validating
                        // GG issue. ***********************
 
                        // prop.setProperty("aubTotalCurr",Integer.toString(AUBTotal[6]));
 
                        // ************** End of Code for comparing grandtotal for
                        // validating GG issue *******************
 
                    } else {
                        System.out.println("AUB Infile not present");
                    }
                } catch (Exception e) {
 
                    System.out.println("Exception in AUB report output...");
                    fin.close();
                    fo.close();
                    System.out.println(e);
                }
 
                // ********************************************************************************************************
 
                FileOutputStream fos = new FileOutputStream(new File(prop.getProperty("output") + timeStamp + ".xls"));
                book.write(fos);
                fos.flush();
                fos.close();
                System.out.println("written ");
 
                text.append("</tr>");
                text.append("</table>");
 
               // SendEmail email = new SendEmail();
              //  email.sendReportMail(text);
 
                // ************************ Pending Code
                // **********************************************
 
                /*
                 * // GG issue validating
                 *
                 * int result = ggIssueChecking();
                 *
                 * if(result==0) { email.sendGGIssue(ggIssue); }
                 */
 
                prop.setProperty("aboTotalPrev", prop.getProperty("aboTotalCurr"));
                prop.setProperty("afsTotalPrev", prop.getProperty("afsTotalCurr"));
                prop.setProperty("aubTotalPrev", prop.getProperty("aubTotalCurr"));
                prop.setProperty("bshTotalPrev", prop.getProperty("bshTotalCurr"));
 
                prop.store(fo, null);
                fo.close();
                fin.close();
    }
 
    /*
     * public static int ggIssueChecking() {
     * if(prop.getProperty("aboTotalCurr").equals(prop.getProperty(
     * "aboTotalPrev"))) { ArrayList<String> at = new ArrayList<>(); at.add(0,
     * prop.getProperty("aboTotalCurr")); at.add(1,
     * prop.getProperty("aboTotalPrev")); ggIssue.put("ABO", at);
     * System.out.println("........"+at.get(0)+".........."+at.get(1)+
     * "........."+ggIssue.get("ABO"));
     *
     * } if(prop.getProperty("afsTotalCurr").equals(prop.getProperty(
     * "afsTotalPrev"))) { ArrayList<String> at = new ArrayList<>(); at.add(0,
     * prop.getProperty("afsTotalCurr")); at.add(1,
     * prop.getProperty("afsTotalPrev")); ggIssue.put("AFS", at);
     * System.out.println("........"+at.get(0)+".........."+at.get(1));
     *
     * } if(prop.getProperty("aubTotalCurr").equals(prop.getProperty(
     * "aubTotalPrev"))) { ArrayList<String> at = new ArrayList<>(); at.add(0,
     * prop.getProperty("aubTotalCurr")); at.add(1,
     * prop.getProperty("aubTotalPrev")); ggIssue.put("AUB", at);
     * System.out.println("........"+at.get(0)+".........."+at.get(1));
     *
     * } if(prop.getProperty("bshTotalCurr").equals(prop.getProperty(
     * "bshTotalPrev"))) { ArrayList<String> at = new ArrayList<>(); at.add(0,
     * prop.getProperty("bshTotalCurr")); at.add(1,
     * prop.getProperty("bshTotalPrev")); ggIssue.put("BSH", at);
     * System.out.println("........"+at.get(0)+".........."+at.get(1));
     *
     * }
     *
     * if(!ggIssue.isEmpty()) { // new
     * SendEmail().abnormalityMailDrafting(reporttext);
     *
     * System.out.println("GG issue started ---------------------");
     * System.out.println(
     * "Environmet \t Previous Grand Total \t Current Grand Total");
     *
     * for(Map.Entry<String,ArrayList<String>> entry : ggIssue.entrySet()) {
     * String key = entry.getKey(); ArrayList<String> value = ggIssue.get(key);
     *
     * System.out.println(key + " => " + value);.get(0)+"\t"+value.get(1)); }
     * return 0;
     *
     * } else { System.out.println("No GG issue"); return 1;
     *
     * }
     *
     *
     * }
     */
}