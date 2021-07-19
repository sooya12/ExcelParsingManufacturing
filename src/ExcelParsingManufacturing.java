import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * 메타데이터 가공 코드
 * 메타데이터의 `특정 cell 값`을 기준으로 특정 cell마다 템플릿에 맞게 excel 파일을 만듦
 *
 * @author sooya12
 */
public class ExcelParsingManufacturing {

    public static void main(String[] args) throws IOException {
        FileInputStream file = new FileInputStream("originalFile/메타데이터.xlsx"); // 메타데이터 파일 경로

        XSSFWorkbook workbook = new XSSFWorkbook(file); // 메타데이터를 읽어온 excel 파일
        XSSFSheet sheet = workbook.getSheetAt(0); // 1번째 시트

        int rowCnt = sheet.getPhysicalNumberOfRows(); // 행 개수
        int cellCnt = 0;
        XSSFRow row;

        List<String> checkId = new ArrayList<>(); // `특정 cell 값` 중복확인

        for (int i = 0; i < rowCnt; i++) { // 행 개수만큼 반복
            row = sheet.getRow(i); // i번째 행

            if(i == 0) { // 첫번째 행이면, 열 개수 구하기 (첫번째 행이 헤더인 경우)
                cellCnt = row.getLastCellNum();
                continue;
            }

            XSSFCell idCell = row.getCell(2); // `특정 cell 값` 열
            String currentId = getCellValue(idCell); // `특정 cell 값` 값

            XSSFCell id2Cell = row.getCell(3); // `특정 cell 값2` 열
            String currentId2 = getCellValue(id2Cell); // `특정 cell 값` 값

            if(!checkId.contains(currentId)) { // 해당 `특정 cell 값`의 파일이 아직 생성되지 않은 경우
                System.out.println(currentId + " " + currentId2);

                checkId.add(currentId); // `특정 cell 값` 중복처리

                XSSFWorkbook newWorkbook = new XSSFWorkbook(new FileInputStream("originalFile/템플릿.xlsx")); // 템플릿 excel 불러오기
                XSSFSheet newSheet = newWorkbook.getSheetAt(0); // 첫번째 시트
                int newRowIdx = 1;

                XSSFRow newRow;
                XSSFCell newCell;

                for (int r = i; r < rowCnt; r++) {
                    row = sheet.getRow(r); // 기존 시트의 r번째 행
                    XSSFCell rIdCell = row.getCell(2); // r번째 행의 `특정 cell 값`

                    String value = getCellValue(rIdCell);

                    if(currentId.equals(value)) { // r번째 행의 `특정 cell 값`이 현재 수집 중인 `특정 cell 값`와 일치하는 경우
                        newRow = newSheet.createRow(newRowIdx++); // 행 추가, 1부터 시작

                        for (int c = 0; c < cellCnt; c++) {
                            XSSFCell insertCell = row.getCell(c);

                            String insertValue = getCellValue(insertCell);

                            newCell = newRow.createCell(c);
                            newCell.setCellValue(insertValue);
                        }
                    }
                }

                String regEx = "([\\\\]|[/]|[:]|[*]|[?]|[\"]|<|>|[|])"; // 파일명에 쓰지 못하는 특수문자 {\, /, :, *, ?, ", <, >, |} 치환
                currentId = currentId.replaceAll(regEx, "");
                currentId2 = currentId2.replaceAll(regEx, "");

                File newFile = new File("storage/" + currentId + "_" + currentId2 + ".xlsx"); // `특정 cell 값`_`특정 cell 값2`.xlsx 파일명 설정
                newFile.createNewFile();

                FileOutputStream fos = new FileOutputStream(newFile);

                newWorkbook.write(fos);

                if(fos != null) {
                    fos.close();
                }
            }
        }
    }

    /**
     * cell 값을 변환하는 메서드
     * @param xssfCell
     * @return
     */
    public static String getCellValue(XSSFCell xssfCell) {
        String cellValue = "";

        if(xssfCell != null) {
            switch (xssfCell.getCellType()) {
                case BOOLEAN:
                    cellValue = xssfCell.getStringCellValue();
                    break;
                case ERROR:
                    cellValue = xssfCell.getErrorCellValue() + "";
                    break;
                case FORMULA:
                    cellValue = xssfCell.getCellFormula();
                    break;
                case NUMERIC:
                    xssfCell.setCellType(CellType.STRING);
                    cellValue = xssfCell.getStringCellValue();
                    break;
                case STRING:
                    cellValue = xssfCell.getStringCellValue();
                    break;
                case _NONE:
                case BLANK:
                default:
                    cellValue = "";
            }
        }

        return cellValue;
    }
}
