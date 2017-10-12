
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;

/**
 * Created by verenich on 20.09.2017.
 */
class Parser {
    public static void main(String[] args) throws IOException {

//        File file=new File("dataMK_EN.xlsx");

        File file=new File("dataGB.xlsx");
        FileInputStream inputStream=new FileInputStream(file);
        XSSFWorkbook workbook=new XSSFWorkbook(inputStream);

//        parse("dataMK1.txt","dataMK1",workbook);
//        parse("dataMK2.txt","dataMK2",workbook);
//        parse("dataMK3.txt","dataMK3",workbook);
//        parse("dataMK5.txt","dataMK5",workbook);

        parse("dataM1.txt","dataM1",workbook);
        parse("dataM2.txt","dataM2",workbook);
        parse("dataM3.txt","dataM3",workbook);
        parse("dataM4.txt","dataM4",workbook);
        parse("dataR1.txt","dataR1",workbook);
        parse("dataR2.txt","dataR2",workbook);
        parse("dataR3.txt","dataR3",workbook);
        parse("dataR4.txt","dataR4",workbook);
        parse("dataR5.txt","dataR5",workbook);
        parse("dataR6.txt","dataR6",workbook);
        parse("dataR7.txt","dataR7",workbook);
        parse("dataR8.txt","dataR8",workbook);
        parse("dataC1.txt","dataC1",workbook);
        parse("dataC2.txt","dataC2",workbook);
        parse("dataC3.txt","dataC3",workbook);
        parse("dataC4.txt","dataC4",workbook);
        parse("dataC5.txt","dataC5",workbook);
        parse("dataPR1.txt","dataPR1",workbook);
        parse("dataPR2.txt","dataPR2",workbook);

        inputStream.close();
        FileOutputStream outputStream=new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();
    }
    public static void parse(String txtName,String sheetName, XSSFWorkbook wb) throws IOException {
        FileInputStream input = new FileInputStream(txtName);
        BufferedReader br = new BufferedReader(new InputStreamReader(input));
        String strLine;
        ArrayList<Double> arrayList = new ArrayList<Double>();
        while ((strLine = br.readLine()) != null) {
            double d = Double.parseDouble(strLine);
            arrayList.add(d);
        }
        input.close();

//        for (Double d : arrayList) {
//            System.out.println(d);
//        }
        XSSFSheet sheet = wb.getSheet(sheetName);
        Iterator<Row> iteratorRow = sheet.rowIterator();
        int count = 0;
        while (iteratorRow.hasNext()) {
            XSSFRow row = (XSSFRow) iteratorRow.next();
            XSSFCell cell = row.getCell(0);
//            cell.setCellValue(arrayList.get(count));
            cell.setCellValue(0);
            System.out.println(cell.toString());
            count++;
        }
        System.out.println("-----------------------");
        wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
    }
}
