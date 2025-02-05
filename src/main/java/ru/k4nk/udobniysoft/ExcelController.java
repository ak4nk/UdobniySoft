package ru.k4nk.udobniysoft;

import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.media.Content;
import io.swagger.v3.oas.annotations.media.Schema;
import io.swagger.v3.oas.annotations.responses.ApiResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;


import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;


@RestController
@Tag(name = "XLSX Service", description = "Find N-th max number in an XLSX file")
public class ExcelController {

    @Operation(
            summary = "Поиск N-го максимального числа в XLSX-файле",
            description = "Метод принимает путь к локальному XLSX-файлу и число N, затем находит N-е максимальное число в указанном файле. "
                          + "Файл должен содержать целые числа в одном столбце. "
                          + "Если N больше количества чисел в файле, возвращается ошибка 404 (NOT FOUND). "
                          + "Если файл пустой, возвращается ошибка 422 (UNPROCESSABLE ENTITY). "
                          + "Если N некорректное (меньше 1), возвращается ошибка 400 (BAD REQUEST).",
            responses = {
                    @ApiResponse(responseCode = "200", description = "Успешный поиск N-го максимального числа",
                            content = @Content(schema = @Schema(type = "integer"))),
                    @ApiResponse(responseCode = "400", description = "Некорректное значение N (меньше 1)"),
                    @ApiResponse(responseCode = "404", description = "Запрошенное число отсутствует в файле"),
                    @ApiResponse(responseCode = "422", description = "Файл пустой или не содержит чисел"),
                    @ApiResponse(responseCode = "500", description = "Ошибка сервера при обработке файла")
            }
    )
    @GetMapping("/find-nth-max")
    public ResponseEntity<Integer> findNthMaxNumber(
            @Parameter(description = "Путь к локальному XLSX-файлу", required = true)
            @RequestParam String filePath,

            @Parameter(description = "N-е максимальное число, которое необходимо найти", required = true, example = "3")
            @RequestParam Integer n) throws IOException {
        if (n < 1) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(null);
        }

        int[] topN = new int[n];
        Arrays.fill(topN, Integer.MIN_VALUE);
        Set<Integer> uniqueNumbers = new HashSet<>();
        int count = 0;

        File file = new File(filePath);
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.NUMERIC) {
                        int num = (int) cell.getNumericCellValue();
                        if (!uniqueNumbers.contains(num)) {
                            uniqueNumbers.add(num);
                            insertIntoTopN(topN, num);
                            count++;
                        }
                    }
                }
            }
        }
        if (count == 0) {
            return ResponseEntity.status(HttpStatus.UNPROCESSABLE_ENTITY).build();
        }
        if (count < n) {
            return ResponseEntity.status(HttpStatus.NOT_FOUND).build();
        }

        return ResponseEntity.ok(topN[n - 1]);
    }

    private void insertIntoTopN(int[] topN, int num) {
        for (int i = 0; i < topN.length; i++) {
            if (num > topN[i]) {
                for (int j = topN.length - 1; j > i; j--) {
                    topN[j] = topN[j - 1];
                }
                topN[i] = num;
                break;
            }
        }
    }
}