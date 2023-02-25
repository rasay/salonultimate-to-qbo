package com.asayclips;

import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.math.MathContext;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * $ java -jar target/salonultimate-to-qbo-1.0-SNAPSHOT.jar {date MM/dd/YYYY} {store #(UT104,UT201 or UT202)}
 */
public class App
{
    private static String _userHome = System.getProperty("user.home");

    public static void main( String[] args )
    {
        if (args.length != 2)
        {
            System.err.println("Usage:  {date MM/dd/YYYY} {store #(UT104,UT201 or UT202)}");
            System.exit(1);
        }
        App app = new App();
        try
        {
            File file = app.findStoreAnalysisReport(args[0], args[1]);
            if (file != null)
                app.readFile(file);
        }
        catch (Exception e)
        {
            System.err.println(e.getMessage());
        }
    }

    class DailyReport
    {
        String date;
        BigDecimal tips;
        BigDecimal salesTax;
        BigDecimal retail;
        BigDecimal total;
        BigDecimal amex;
        BigDecimal discover;
        BigDecimal cash;

        public BigDecimal getBoa()
        {
            return total.subtract(amex).subtract(cash);
        }

        public BigDecimal getService()
        {
            return getBoa().subtract(retail).subtract(salesTax).subtract(tips);
        }

        public String toString()
        {
            return "Date: " + date + ", tips: $" + tips + ", tax: $" + salesTax + ", retail: $" + retail + ", BoA: $" + getBoa() + ", Amex: $" + amex;
        }
    }

    public void readFile(File file)
    {
        try
        {
            List<DailyReport> reports = new ArrayList<DailyReport>();

            FileInputStream fileInputStream = new FileInputStream(file);
            HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
            HSSFSheet sheet = workbook.getSheet("Worksheet");

            for (int i=0; i<6; i++)
                reports.add(getDailyReport(sheet, i));

            printReports(reports);
        }
        catch (Exception e)
        {
            System.out.println("ERROR : " + e.toString());
            e.printStackTrace();
        }
    }

    private static SimpleDateFormat _shortDateFormat = new SimpleDateFormat("MM/dd/yyyy");

    private File findStoreAnalysisReport(String date, String storeNumber) throws Exception
    {
        File file = null;
        String targetDate = _shortDateFormat.format(_shortDateFormat.parse(date));

        File dir = new File(_userHome + "/Downloads/");
        for (File f : dir.listFiles())
        {
            if (f.getName().startsWith("Store_Analysis")
                    && f.getName().endsWith(".xls")
                    && isTheRightStoreAnalysisFile(f, targetDate, storeNumber))
            {
                System.out.printf("Found store analysis report for store: %s (%s) %s%n",
                        storeNumber, date, f.getAbsolutePath());
                file = f;
                break;
            }
        }
        if (file == null)
            System.out.printf("Download store analysis report for store: %s (%s)%n",
                    storeNumber, date);

        return file;
    }

    private boolean isTheRightStoreAnalysisFile(File file, String date, String storeNumber) throws Exception
    {
        FileInputStream inputStream = new FileInputStream(file);
        try
        {
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            HSSFSheet sheet = workbook.getSheet("Worksheet");

            HSSFRow row = sheet.getRow(0);
            if (!storeNumber.toLowerCase().equals(row.getCell(1).toString().toLowerCase()))
                return false;

            row = sheet.getRow(1);

            for (int i=1; i<8; i++)
                if (date.equals(row.getCell(i).getStringCellValue()))
                    return true;

            return false;
        }
        finally
        {
            inputStream.close();
        }
    }

    private DailyReport getDailyReport(HSSFSheet sheet, int dayOfWeek)
    {
        DailyReport report = new DailyReport();

        // get date
        HSSFRow row = sheet.getRow(5);
        report.date = row.getCell(dayOfWeek+2).getStringCellValue();

        // get tips
        row = sheet.getRow(33);
        report.tips = new BigDecimal(row.getCell(dayOfWeek+2).getNumericCellValue());

        // get sales tax
        row = sheet.getRow(9);
        report.salesTax = new BigDecimal(row.getCell(dayOfWeek+2).getNumericCellValue(), MathContext.DECIMAL32);

        // get retail
        row = sheet.getRow(7);
        report.retail = new BigDecimal(row.getCell(dayOfWeek+2).getNumericCellValue(), MathContext.DECIMAL32);

        // get total
        row = sheet.getRow(28);
        report.total = new BigDecimal(row.getCell(dayOfWeek+2).getNumericCellValue(), MathContext.DECIMAL32);

        // get discover
        row = sheet.getRow(27);
        report.discover = new BigDecimal(row.getCell(dayOfWeek+2).getNumericCellValue(), MathContext.DECIMAL32);

        // get cash
        row = sheet.getRow(23);
        report.cash = new BigDecimal(row.getCell(dayOfWeek+2).getNumericCellValue(), MathContext.DECIMAL32);

        // get Amex
        row = sheet.getRow(26);
        report.amex = new BigDecimal(row.getCell(dayOfWeek+2).getNumericCellValue(), MathContext.DECIMAL32);

        return report;
    }

    private void printReports(List<DailyReport> reports)
    {
        for (int i=0; i<6; i++) {
            System.out.println("*******************************************************");
            System.out.println(reports.get(i).date);
            System.out.printf("Tips:      %10.2f\n", reports.get(i).tips);
            System.out.printf("Sales Tax: %10.2f\n", reports.get(i).salesTax);
            System.out.printf("Retail:    %10.2f\n", reports.get(i).retail);
            System.out.printf("Service:   %10.2f\n", reports.get(i).getService());
            System.out.printf("BoA:       %10.2f\n", reports.get(i).getBoa());
            System.out.printf("Amex:      %10.2f\n", reports.get(i).amex);
        }
    }
}
