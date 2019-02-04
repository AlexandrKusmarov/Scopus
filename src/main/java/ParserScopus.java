import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * This programm works only with the site Scopus.com.
 * It is also necessary to have an account with all the rights and log in.
 * How it works you can read in Readme.txt file.
 */

public class ParserScopus {
    public static void main(String[] args) throws IOException {
        String url = args[0];
        String urlBibl = args[1];
        String xmlPath;
        File input = new File(url);
        Document doc = Jsoup.parse(input, "utf-8", url);
        Document docBibl = Jsoup.parse(new File(urlBibl), "utf-8", urlBibl);
        Elements elements = doc.getElementsByTag("tbody");
        String authorName = doc.getElementsByClass("wordBreakWord").text();
        Elements hIndexAndDocs = doc.getElementsByClass("FontLarge");
        Elements rows = elements.select("tr[class=searchArea]");
        Elements elemDocBibl = docBibl.select("p");
        Elements span = doc.select(".secondaryLink .anchorText");
        Elements authID = doc.select(".authId");

        xmlPath = url.replace(".htm", ".xls");

        File excellFile = new File(xmlPath);
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Publications");

        HSSFRow row1 = sheet.createRow(0);
        HSSFCell name1 = row1.createCell(0);
        name1.setCellValue("Authors");
        HSSFCell hind = row1.createCell(1);
        hind.setCellValue("H-index");
        HSSFCell authdocs = row1.createCell(2);
        authdocs.setCellValue("Authors docs");
        HSSFCell ttlnumb = row1.createCell(4);
        ttlnumb.setCellValue("Total number of citation");
        HSSFCell id = row1.createCell(3);
        id.setCellValue("Author ID");
        HSSFCell orcid = row1.createCell(5);
        orcid.setCellValue("ORCID");

        HSSFRow row = sheet.createRow(1);
        HSSFCell name = row.createCell(0);
        name.setCellValue(authorName);
        HSSFCell hInd = row.createCell(1);
        hInd.setCellValue(hIndexAndDocs.first().text());
        HSSFCell authDoc = row.createCell(2);
        authDoc.setCellValue(hIndexAndDocs.get(1).text());
        HSSFCell totcit = row.createCell(4);
        totcit.setCellValue(hIndexAndDocs.get(2).text());
        HSSFCell idValue = row.createCell(3);
        Pattern p = Pattern.compile("[\\d]*[\\d]");
        Matcher m = p.matcher(authID.text());
        while (m.find()) {
            idValue.setCellValue(m.group());
        }
        HSSFCell arcidValue = row.createCell(5);
        if (span.get(7).text().contains("http://orcid.org/")) {
            arcidValue.setCellValue(span.get(7).text().substring(17));
        }

        HSSFRow row3 = sheet.createRow(3);
        HSSFCell docname = row3.createCell(0);
        docname.setCellValue("Document name");
        HSSFCell aut = row3.createCell(1);
        aut.setCellValue("Authors");
        HSSFCell year = row3.createCell(2);
        year.setCellValue("Year");
        HSSFCell source = row3.createCell(3);
        source.setCellValue("Source");
        HSSFCell citations = row3.createCell(4);
        citations.setCellValue("Citations");
        HSSFCell bibl = row3.createCell(5);
        bibl.setCellValue("Bibliographic description");

        int indexForXLS = 4;
        int indexBibl = 2;
        for (Element searchArea : rows) {
            HSSFRow rowww = sheet.createRow(indexForXLS);
            HSSFCell cell = rowww.createCell(0);
            cell.setCellValue(searchArea.select(".ddmDocTitle").text());
            HSSFCell cell1 = rowww.createCell(1);
            cell1.setCellValue(searchArea.select(".ddmAuthorList").text());
            HSSFCell cell2 = rowww.createCell(2);
            cell2.setCellValue(searchArea.select(".ddmPubYr").text());
            HSSFCell cell3 = rowww.createCell(3);
            cell3.setCellValue(searchArea.child(4).text());
            HSSFCell cell4 = rowww.createCell(4);
            cell4.setCellValue(searchArea.children().last().text());
            HSSFCell cell5 = rowww.createCell(5);
            if (sheet.getLastRowNum() >= indexBibl) {
                cell5.setCellValue(elemDocBibl.get(indexBibl).text()
                        .substring(elemDocBibl.get(indexBibl).text().indexOf(".") + 2)
                        .replace("Available from: www.scopus.com", ""));
                indexBibl++;
            } else break;
            indexForXLS++;
        }
        workbook.write(new FileOutputStream(excellFile));
    }
}
