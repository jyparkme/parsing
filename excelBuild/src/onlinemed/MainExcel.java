package onlinemed;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.util.Set;

public class MainExcel {
    public static void main(String[] args) throws Exception {

        String path = "C:\\Onlinemed";
        File folder = new File(path);
        if (!folder.exists()) {
            folder.mkdir();
            System.out.println("온라인메드폴더가 생성되었습니다.");
        }

        String medPath = "C:\\Onlinemed\\원본(" + LocalDate.now().minusDays(1) + ")";
        File medFolder = new File(medPath);
        if (!medFolder.exists()) {
            medFolder.mkdir();
            System.out.println(LocalDate.now().minusDays(1) + " 원본폴더가 생성되었습니다.");
        } else {System.out.println("폴더가 이미 존재합니다.");}

        String KMIPath = "C:\\Onlinemed\\KMI(" + LocalDate.now().minusDays(1) + ")";
        File KMIFolder = new File(KMIPath);
        if (!KMIFolder.exists()) {
            KMIFolder.mkdir();
            System.out.println(LocalDate.now().minusDays(1) + " KMI폴더가 생성되었습니다.");
        } else {System.out.println("폴더가 이미 존재합니다.");}

        ExcelDown excelDown = new ExcelDown();
        for (int i = 0; i < 7; i++) {
            excelDown.idNum = i;
            excelDown.id = excelDown.findId(excelDown.idNum);
            excelDown.getFileDown();
        }


        for (int i=0; i<7; i++) {
            ExcelBuild excel = new ExcelBuild();
            excel.setIds();
            excel.setExcel(i);
            String excelName = excel.fileName + ".xlsx";

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet(excel.fileName);
            XSSFRow row = null;
            XSSFCell cell = null;
            int rowNum = 0;
            int cellTitle = 0;
            Set<Integer> keySet = excel.mList.keySet();

            row = sheet.createRow(rowNum++);
            for (int m=0; m<121; m++) {
                cell = row.createCell(cellTitle++);
                cell.setCellValue(m);
            }

            for (Integer key : keySet) {
                row = sheet.createRow(rowNum++);
                String[] lists = excel.mList.get(key);
                int cellNum = 0;
                for (String list : lists) {
                    cell = row.createCell(cellNum++);
                    cell.setCellValue(list);
                }
            }

            FileOutputStream out = new FileOutputStream(new File(KMIPath, excelName));
            workbook.write(out);
            out.close();
            workbook.close();
        }

        System.out.println("다운로드가 완료되었습니다.");
    }
}
