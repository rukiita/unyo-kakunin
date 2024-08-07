package com.example;
import java.io.File;
import java.io.IOException;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

public class EmployeeCheckList2 {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        // ユーザーから月を入力してもらう
        System.out.print("月を入力してください: ");
        int month = scanner.nextInt();

        // グループの選択肢を表示
        System.out.println("グループを選択してください:");
        System.out.println("1: 販売");
        System.out.println("2: 広告");
        System.out.println("3: 管理");
        System.out.println("4: 営業");
        System.out.println("5: 全て");
        int groupChoice = scanner.nextInt();
        scanner.close();

        // 指定された月に基づいてチェックする列を決定
        List<String> columnsToCheck = determineColumnsToCheck(month);

        try {
            // Excelファイルを読み込む
            Workbook workbook = WorkbookFactory.create(new File("path/管理ファイル"));
            Sheet mainSheet = workbook.getSheetAt(0);
            Sheet nameSheet = workbook.getSheet("名簿");

            // 名前リストとグループリストを取得
            Map<Integer, String> employeeNames = new HashMap<>();
            Map<Integer, String> employeeGroups = new HashMap<>();
            for (int rowNumber = 2; rowNumber <= 86; rowNumber++) {
                Row row = nameSheet.getRow(rowNumber - 1);
                if (row == null) continue;

                Cell nameCell = row.getCell(2);
                Cell groupCell = row.getCell(6);
                if (nameCell == null || nameCell.getCellType() != CellType.STRING || groupCell == null || groupCell.getCellType() != CellType.STRING) continue;

                // ここで対応する行番号を正確に計算します
                int correspondingRowNumber = rowNumber + 14;
                employeeNames.put(correspondingRowNumber, nameCell.getStringCellValue());
                employeeGroups.put(correspondingRowNumber, groupCell.getStringCellValue());
            }

            // メインシートの16行目から102行目までの各社員について処理
            for (int rowNumber = 16; rowNumber <= 102; rowNumber++) {
                Row row = mainSheet.getRow(rowNumber - 1); // 行番号は0ベース
                if (row == null) continue;

                // 正確な対応を確認します
                String employeeName = employeeNames.get(rowNumber);
                String employeeGroup = employeeGroups.get(rowNumber);
                if (employeeName == null || employeeGroup == null) continue;

                // グループフィルタリング
                if (!isGroupMatch(groupChoice, employeeGroup)) continue;

                List<Integer> uncompletedItems = new ArrayList<>();

                // 指定された列についてチェック
                for (String column : columnsToCheck) {
                    int cellIndex = CellReference.convertColStringToIndex(column);
                    Cell cell = row.getCell(cellIndex);
                    if (cell == null || cell.getCellType() != CellType.STRING) continue;

                    if ("対象".equals(cell.getStringCellValue())) {
                        int checkNumber = getCheckNumber(column, month % 2 == 0); // 偶数月かどうかでチェック番号を決定
                        uncompletedItems.add(checkNumber);
                    }
                }

                // 未完了のチェック項目があれば出力
                if (!uncompletedItems.isEmpty()) {
                    System.out.print(employeeName + " : ");
                    for (int item : uncompletedItems) {
                        System.out.print(item + " ");
                    }
                    System.out.println();
                }
            }

            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // グループが一致するかどうかを確認するメソッド
    private static boolean isGroupMatch(int groupChoice, String employeeGroup) {
        switch (groupChoice) {
            case 1:
                return "販売".equals(employeeGroup);
            case 2:
                return "広告".equals(employeeGroup);
            case 3:
                return "管理".equals(employeeGroup);
            case 4:
                return "営業".equals(employeeGroup);
            case 5:
                return true; // 全ての場合
            default:
                return false;
        }
    }

    // 月に基づいてチェックする列を決定するメソッド
    private static List<String> determineColumnsToCheck(int month) {
        List<String> columnsToCheck = new ArrayList<>();

        if (month % 2 == 0) { // 偶数月
            columnsToCheck.addAll(Arrays.asList("H", "I", "J", "K", "L", "M", "N", "R", "S", "T"));
            if (month == 6) columnsToCheck.add("O");
            if (month == 4) columnsToCheck.add("P");
            if (month == 10) columnsToCheck.add("Q");
        } else { // 奇数月
            columnsToCheck.addAll(Arrays.asList("J", "L"));
        }

        return columnsToCheck;
    }

    // 列と偶数月かどうかに基づいてチェック番号を取得するメソッド
    private static int getCheckNumber(String column, boolean isEvenMonth) {
        Map<String, Integer> evenMonthCheckNumbers = new HashMap<>();
        evenMonthCheckNumbers.put("H", 1);
        evenMonthCheckNumbers.put("I", 2);
        evenMonthCheckNumbers.put("J", 3);
        evenMonthCheckNumbers.put("K", 4);
        evenMonthCheckNumbers.put("L", 5);
        evenMonthCheckNumbers.put("M", 6);
        evenMonthCheckNumbers.put("N", 7);
        evenMonthCheckNumbers.put("O", 8);
        evenMonthCheckNumbers.put("P", 9);
        evenMonthCheckNumbers.put("Q", 10);
        evenMonthCheckNumbers.put("R", 11);
        evenMonthCheckNumbers.put("S", 12);
        evenMonthCheckNumbers.put("T", 13);

        Map<String, Integer> oddMonthCheckNumbers = new HashMap<>();
        oddMonthCheckNumbers.put("J", 3);
        oddMonthCheckNumbers.put("L", 5);

        return isEvenMonth ? evenMonthCheckNumbers.get(column) : oddMonthCheckNumbers.get(column);
    }
}