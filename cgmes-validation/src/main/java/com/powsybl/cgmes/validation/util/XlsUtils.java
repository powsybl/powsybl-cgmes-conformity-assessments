/**
 * Copyright (c) 2020, RTE (http://www.rte-france.com)
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */
package com.powsybl.cgmes.validation.util;

import com.powsybl.cgmes.validation.CgmesValidationTool;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;

import java.util.Map;

/**
 * @author Jérémy LABOUS <jlabous at silicom.fr>
 */
public final class XlsUtils {

    public static final String V = "v";
    public static final String ANGLE = "angle";
    private static double[] diffAngle = new double[2];
    private static int[] rowNum = new int[1];

    public static double[] getDiffAngle() {
        return diffAngle;
    }

    public static void generateResults(SXSSFWorkbook wb, Map<String, Map<String, String>> results) {
        XSSFSheet sheet = wb.getXSSFWorkbook().getSheet("Results");

        AreaReference ref = wb.getCreationHelper().createAreaReference(
                new CellReference(0, 0), new CellReference(results.entrySet().size(), 5)
        );
        XSSFTable table = sheet.createTable(ref);
        int i = 2;
        for (CTTableColumn column : table.getCTTable().getTableColumns().getTableColumnList()) {
            column.setId(i++);
        }
        // For now, create the initial style in a low-level way
        table.getCTTable().addNewTableStyleInfo();
        table.getCTTable().getTableStyleInfo().setName("TableStyleMedium2");
        table.getCTTable().getTableStyleInfo().setShowRowStripes(true);
        table.getCTTable().addNewAutoFilter().setRef(ref.formatAsString());

        XSSFSheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
        XSSFConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingColorScaleRule();
        XSSFColorScaleFormatting clrFmt = rule1.getColorScaleFormatting();
        clrFmt.getThresholds()[0].setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
        clrFmt.getThresholds()[0].setValue(0d);
        clrFmt.getColors()[0].setARGBHex("FF00FF00");
        clrFmt.getThresholds()[1].setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
        clrFmt.getThresholds()[1].setValue(15d);
        clrFmt.getColors()[1].setARGBHex("FFFFFF00");
        clrFmt.getThresholds()[2].setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
        clrFmt.getThresholds()[2].setValue(25d);
        clrFmt.getColors()[2].setARGBHex("FFFF0000");

        CellRangeAddress[] regions = {CellRangeAddress.valueOf("E2:F" + results.entrySet().size() + 1)};
        sheetCF.addConditionalFormatting(regions, rule1);

        int rowNum = 0;
        int cellNum = 0;
        XSSFRow currentRow = sheet.createRow(rowNum++);
        XSSFCell cell = currentRow.createCell(cellNum++);
        cell.setCellValue("Case file");
        cell = currentRow.createCell(cellNum++);
        cell.setCellValue("Topology");
        cell = currentRow.createCell(cellNum++);
        cell.setCellValue("Import status");
        cell = currentRow.createCell(cellNum++);
        cell.setCellValue("PowerFlow status");
        cell = currentRow.createCell(cellNum++);
        cell.setCellValue("SSH max diff V");
        cell = currentRow.createCell(cellNum);
        cell.setCellValue("SV max diff V");
        for (Map.Entry<String, Map<String, String>> result : results.entrySet()) {
            cellNum = 0;
            currentRow = sheet.createRow(rowNum++);
            cell = currentRow.createCell(cellNum++);
            cell.setCellValue(result.getValue().get(CgmesValidationTool.CASEFILE_KEY));
            cell = currentRow.createCell(cellNum++);
            cell.setCellValue(result.getValue().get(CgmesValidationTool.TOPOLOGY_KEY));
            cell = currentRow.createCell(cellNum++);
            cell.setCellValue(result.getValue().getOrDefault(CgmesValidationTool.ERROR_KEY, "OK"));
            cell = currentRow.createCell(cellNum++);
            cell.setCellValue(result.getValue().get(CgmesValidationTool.LOADFLOW_KEY));
            cell = currentRow.createCell(cellNum++);
            cell.setCellFormula("MAX('" + result.getValue().get(CgmesValidationTool.PROVIDER_KEY) + "'!Q:Q)");
            cell = currentRow.createCell(cellNum);
            cell.setCellFormula("MAX('" + result.getValue().get(CgmesValidationTool.PROVIDER_KEY) + "'!S:S)");
        }
    }

    public static void writeCompareSVResults(SXSSFWorkbook wb, Map<String, SvUtils.Quintuplet> results, String provider) {
        diffAngle[0] = 0;
        diffAngle[1] = 0;
        rowNum[0] = 0;
        XSSFSheet sheet = wb.getXSSFWorkbook().createSheet(provider);
        XSSFTable table;
        writeHeader(sheet);
        writeTableLines(sheet, results);

        AreaReference ref = wb.getCreationHelper().createAreaReference(
                new CellReference(0, 0), new CellReference(rowNum[0] - 1, 22)
        );
        table = sheet.createTable(ref);
        table.setDisplayName(provider);
        table.setName(provider);
        int i = 2;
        for (CTTableColumn column : table.getCTTable().getTableColumns().getTableColumnList()) {
            column.setId(i++);
        }
        table.getCTTable().addNewTableStyleInfo();
        table.getCTTable().getTableStyleInfo().setName("TableStyleMedium2");
        table.getCTTable().getTableStyleInfo().setShowRowStripes(true);
    }

    private static void writeHeader(XSSFSheet sheet) {
        int colNum = 0;
        XSSFRow row = sheet.createRow(rowNum[0]++);
        row.createCell(colNum++).setCellValue("Topological Node");
        row.createCell(colNum++).setCellValue("VL ID");
        row.createCell(colNum++).setCellValue("IIDM bus ID");
        row.createCell(colNum++).setCellValue("Equipment ID");
        row.createCell(colNum++).setCellValue("numcnx");
        row.createCell(colNum++).setCellValue(V);
        row.createCell(colNum++).setCellValue("v_after_import_SSH");
        row.createCell(colNum++).setCellValue("v_after_lf_SSH");
        row.createCell(colNum++).setCellValue("v_after_import_SV");
        row.createCell(colNum++).setCellValue("v_after_lf_SV");
        row.createCell(colNum++).setCellValue(ANGLE);
        row.createCell(colNum++).setCellValue("angle_after_import_SSH");
        row.createCell(colNum++).setCellValue("angle_after_lf_SSH");
        row.createCell(colNum++).setCellValue("angle_after_import_SV");
        row.createCell(colNum++).setCellValue("angle_after_lf_SV");
        row.createCell(colNum++).setCellValue("diff_v_after_import_SSH");
        row.createCell(colNum++).setCellValue("diff_v_after_lf_SSH");
        row.createCell(colNum++).setCellValue("diff_v_after_import_SV");
        row.createCell(colNum++).setCellValue("diff_v_after_lf_SV");
        row.createCell(colNum++).setCellValue("diff_angle_after_import_SSH");
        row.createCell(colNum++).setCellValue("diff_angle_after_lf_SSH");
        row.createCell(colNum++).setCellValue("diff_angle_after_import_SV");
        row.createCell(colNum).setCellValue("diff_angle_after_lf_SV");
    }

    private static void writeTableLines(XSSFSheet sheet, Map<String, SvUtils.Quintuplet> results) {
        results.forEach((id, quintuplet) -> {
            XSSFRow currentRow = sheet.createRow(rowNum[0]++);
            int cellNum = 0;
            currentRow.createCell(cellNum++).setCellValue(id);
            currentRow.createCell(cellNum++).setCellValue(quintuplet.voltageLevelId);
            currentRow.createCell(cellNum++).setCellValue(quintuplet.busId);
            currentRow.createCell(cellNum++).setCellValue(quintuplet.equipmentId);
            currentRow.createCell(cellNum++).setCellValue(quintuplet.connectedComponentNumber);
            if (Double.isNaN(quintuplet.v1)) {
                cellNum++;
            } else {
                currentRow.createCell(cellNum++).setCellValue(quintuplet.v1);
            }
            if (Double.isNaN(quintuplet.v2)) {
                cellNum++;
            } else {
                currentRow.createCell(cellNum++).setCellValue(quintuplet.v2);
            }
            if (Double.isNaN(quintuplet.v3)) {
                cellNum++;
            } else {
                currentRow.createCell(cellNum++).setCellValue(quintuplet.v3);
            }
            if (Double.isNaN(quintuplet.v4)) {
                cellNum++;
            } else {
                currentRow.createCell(cellNum++).setCellValue(quintuplet.v4);
            }
            if (Double.isNaN(quintuplet.v5)) {
                cellNum++;
            } else {
                currentRow.createCell(cellNum++).setCellValue(quintuplet.v5);
            }
            if (Double.isNaN(quintuplet.angle1)) {
                cellNum++;
            } else {
                currentRow.createCell(cellNum++).setCellValue(quintuplet.angle1);
            }
            if (Double.isNaN(quintuplet.angle2)) {
                cellNum++;
            } else {
                currentRow.createCell(cellNum++).setCellValue(quintuplet.angle2);
            }
            if (Double.isNaN(quintuplet.angle3)) {
                cellNum++;
            } else {
                currentRow.createCell(cellNum++).setCellValue(quintuplet.angle3);
            }
            if (Double.isNaN(quintuplet.angle4)) {
                cellNum++;
            } else {
                currentRow.createCell(cellNum++).setCellValue(quintuplet.angle4);
            }
            if (Double.isNaN(quintuplet.angle5)) {
                cellNum++;
            } else {
                currentRow.createCell(cellNum++).setCellValue(quintuplet.angle5);
            }
            if (!Double.isNaN(quintuplet.v1) && !Double.isNaN(quintuplet.v2)) {
                currentRow.createCell(cellNum++).setCellValue(Math.abs(quintuplet.v2 - quintuplet.v1));
            } else {
                cellNum++;
            }
            if (!Double.isNaN(quintuplet.v1) && !Double.isNaN(quintuplet.v3)) {
                currentRow.createCell(cellNum++).setCellValue(Math.abs(quintuplet.v3 - quintuplet.v1));
            } else {
                cellNum++;
            }
            if (!Double.isNaN(quintuplet.v1) && !Double.isNaN(quintuplet.v4)) {
                currentRow.createCell(cellNum++).setCellValue(Math.abs(quintuplet.v4 - quintuplet.v1));
            } else {
                cellNum++;
            }
            if (!Double.isNaN(quintuplet.v1) && !Double.isNaN(quintuplet.v5)) {
                currentRow.createCell(cellNum++).setCellValue(Math.abs(quintuplet.v5 - quintuplet.v1));
            } else {
                cellNum++;
            }

            if (!Double.isNaN(quintuplet.angle1) && !Double.isNaN(quintuplet.angle2)) {
                currentRow.createCell(cellNum++).setCellValue(Math.abs(quintuplet.angle2 - quintuplet.angle1));
            } else {
                cellNum++;
            }
            if (!Double.isNaN(quintuplet.angle1) && !Double.isNaN(quintuplet.angle3)) {
                currentRow.createCell(cellNum++).setCellValue(Math.abs(quintuplet.angle3 - quintuplet.angle1 - diffAngle[0]));
            } else {
                cellNum++;
            }
            if (!Double.isNaN(quintuplet.angle1) && !Double.isNaN(quintuplet.angle4)) {
                currentRow.createCell(cellNum++).setCellValue(Math.abs(quintuplet.angle4 - quintuplet.angle1));
            } else {
                cellNum++;
            }
            if (!Double.isNaN(quintuplet.angle1) && !Double.isNaN(quintuplet.angle5)) {
                currentRow.createCell(cellNum).setCellValue(Math.abs(quintuplet.angle5 - quintuplet.angle1 - diffAngle[1]));
            }
        });
    }

    private XlsUtils() {
    }
}