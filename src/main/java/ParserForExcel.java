import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

public class ParserForExcel {
    static int i = 2;
    private static Document getDocument()throws IOException{
        String url = "https://www.avito.ru/respublika_krym/avtomobili?pmax=50000&pmin=0";
        Document page = Jsoup.parse(new URL(url),10000);
        return page;
    }

    private static HSSFCellStyle createStyleForTitle(HSSFWorkbook workbook){
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        return style;
    }
    public static List<Employee>employeeListDAO(){
        List<Employee>list = new ArrayList<Employee>();
        return list;
    }

    public static void main(String[] args) throws IOException {
        List<Employee>list = new ArrayList<Employee>();
        Document page = getDocument();
        Elements elementes = page.select("div[class = item item_table clearfix js-catalog-item-enum  item-with-contact js-item-extended]");
        System.out.println("Все машины Крыма которые дешевле 50к рублей!!!");
        for(Element elements1:elementes){
        Element model = elements1.selectFirst("span[itemprop = name]");
        Element price = elements1.selectFirst("span[class= price]");
        Element info = elements1.selectFirst("div[class = specific-params specific-params_block]");
        Element date = elements1.selectFirst("div[class = data]");
        Employee employee = new Employee(model.text(),price.text(),info.text(),date.text());
        list.add(employee);
        i++;
    }
        Document page2 = Jsoup.parse(new URL("https://www.avito.ru/respublika_krym/avtomobili?p=2&cd=1&pmax=50000&pmin=0"),50000);
        Elements elementes2 = page2.select("div[class = item item_table clearfix js-catalog-item-enum  item-with-contact js-item-extended]");
        for(Element elements1:elementes2){
            Element model = elements1.selectFirst("span[itemprop = name]");
            Element price = elements1.selectFirst("span[class= price]");
            Element info = elements1.selectFirst("div[class = specific-params specific-params_block]");
            Element date = elements1.selectFirst("div[class = data]");
            Employee employee = new Employee(model.text(),price.text(),info.text(),date.text());
            list.add(employee);
            i++;
        }
        Document page3 = Jsoup.parse(new URL("https://www.avito.ru/respublika_krym/avtomobili?p=3&cd=1&pmax=50000&pmin=0"),50000);
        Elements elementes3 = page3.select("div[class = item item_table clearfix js-catalog-item-enum  item-with-contact js-item-extended]");
        for(Element elements1:elementes3){
            Element model = elements1.selectFirst("span[itemprop = name]");
            Element price = elements1.selectFirst("span[class= price]");
            Element info = elements1.selectFirst("div[class = specific-params specific-params_block]");
            Element date = elements1.selectFirst("div[class = data]");
            Employee employee = new Employee(model.text(),price.text(),info.text(),date.text());
            list.add(employee);
            i++;
        }
        Document page4 = Jsoup.parse(new URL("https://www.avito.ru/respublika_krym/avtomobili?p=4&cd=1&pmax=50000&pmin=0"),50000);
        Elements elementes4 = page4.select("div[class = item item_table clearfix js-catalog-item-enum  item-with-contact js-item-extended]");
        for(Element elements1:elementes4){
            Element model = elements1.selectFirst("span[itemprop = name]");
            Element price = elements1.selectFirst("span[class= price]");
            Element info = elements1.selectFirst("div[class = specific-params specific-params_block]");
            Element date = elements1.selectFirst("div[class = data]");
            Employee employee = new Employee(model.text(),price.text(),info.text(),date.text());
            list.add(employee);
            i++;
        }
        Document page5 = Jsoup.parse(new URL("https://www.avito.ru/respublika_krym/avtomobili?p=5&cd=1&pmax=50000&pmin=0"),50000);
        Elements elementes5 = page5.select("div[class = item item_table clearfix js-catalog-item-enum  item-with-contact js-item-extended]");
        for(Element elements1:elementes5){
            Element model = elements1.selectFirst("span[itemprop = name]");
            Element price = elements1.selectFirst("span[class= price]");
            Element info = elements1.selectFirst("div[class = specific-params specific-params_block]");
            Element date = elements1.selectFirst("div[class = data]");
            Employee employee = new Employee(model.text(),price.text(),info.text(),date.text());
            list.add(employee);
            i++;
        }
        Document page6 = Jsoup.parse(new URL("https://www.avito.ru/respublika_krym/avtomobili?p=6&cd=1&pmax=50000&pmin=0"),50000);
        Elements elementes6 = page6.select("div[class = item item_table clearfix js-catalog-item-enum  item-with-contact js-item-extended]");
        for(Element elements1:elementes6){
            Element model = elements1.selectFirst("span[itemprop = name]");
            Element price = elements1.selectFirst("span[class= price]");
            Element info = elements1.selectFirst("div[class = specific-params specific-params_block]");
            Element date = elements1.selectFirst("div[class = data]");
            Employee employee = new Employee(model.text(),price.text(),info.text(),date.text());
            list.add(employee);
            i++;
        }
        Document page7 = Jsoup.parse(new URL("https://www.avito.ru/respublika_krym/avtomobili?p=7&cd=1&pmax=50000&pmin=0"),50000);
        Elements elementes7 = page7.select("div[class = item item_table clearfix js-catalog-item-enum  item-with-contact js-item-extended]");
        for(Element elements1:elementes7){
            Element model = elements1.selectFirst("span[itemprop = name]");
            Element price = elements1.selectFirst("span[class= price]");
            Element info = elements1.selectFirst("div[class = specific-params specific-params_block]");
            Element date = elements1.selectFirst("div[class = data]");
            Employee employee = new Employee(model.text(),price.text(),info.text(),date.text());
            list.add(employee);
            i++;
        }
        HSSFWorkbook workbook = new HSSFWorkbook();

        HSSFSheet sheet = workbook.createSheet();

        int rowNum = 0;
        Cell cell ;
        Row row;
        HSSFCellStyle style = createStyleForTitle(workbook);
        row = sheet.createRow(rowNum);
        cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Модель");
        cell.setCellStyle(style);
        cell = row.createCell(1, CellType.STRING);
        cell.setCellValue("Цена");
        cell.setCellStyle(style);
        cell = row.createCell(2, CellType.STRING);
        cell.setCellValue("Информация");
        cell.setCellStyle(style);
        cell = row.createCell(3, CellType.STRING);
        cell.setCellValue("Дата");
        cell.setCellStyle(style);
        for (Employee employee: list){
            rowNum++;
            row = sheet.createRow(rowNum);
            cell = row.createCell(0,CellType.STRING);
            cell.setCellValue(employee.getModel());
            cell = row.createCell(1,CellType.STRING);
            cell.setCellValue(employee.getPrice());
            cell = row.createCell(2,CellType.STRING);
            cell.setCellValue(employee.getInfo());
            cell = row.createCell(3,CellType.STRING);
            cell.setCellValue(employee.getDate());
        }
        File file = new File("C://pezda//ParseForExcel.xls");
        file.getParentFile().mkdirs();
        FileOutputStream stream = new FileOutputStream(file);
        workbook.write(stream);
        System.out.println("Created file "+ file.getAbsolutePath());
    }
}
