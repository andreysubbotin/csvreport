package com.company.csvreport.core;

import com.haulmont.cuba.core.global.AppBeans;
import com.haulmont.yarg.exception.ReportingException;
import com.haulmont.yarg.formatters.CustomReport;
import com.haulmont.yarg.formatters.ReportFormatter;
import com.haulmont.yarg.formatters.factory.FormatterFactoryInput;
import com.haulmont.yarg.formatters.factory.ReportFormatterFactory;
import com.haulmont.yarg.structure.BandData;
import com.haulmont.yarg.structure.Report;
import com.haulmont.yarg.structure.ReportOutputType;
import com.haulmont.yarg.structure.ReportTemplate;
import org.apache.commons.io.IOUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.xlsx4j.exceptions.Xlsx4jException;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.SheetData;
import org.xlsx4j.sml.Worksheet;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class CsvReport implements CustomReport {

    private static final char DEFAULT_SEPARATOR = ',';

    @Override
    public byte[] createReport(Report report, BandData rootBand, Map<String, Object> params) {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        ReportFormatter xslxFormatter = createXslxFormatter(report, rootBand, outputStream);
        xslxFormatter.renderDocument();
        return convertXslxToCsv(outputStream.toByteArray());
    }

    protected ReportFormatter createXslxFormatter(Report report, BandData rootBand, OutputStream outputStream) {
        ReportTemplate reportTemplate = report.getReportTemplates().get("DEFAULT");
        ReportFormatterFactory formatterFactory = AppBeans.get("reporting_lib_FormatterFactory");
        FormatterFactoryInput formatterFactoryInput = new FormatterFactoryInput("xlsx", rootBand, new ReportTemplateWrapper(reportTemplate), outputStream);
        return formatterFactory.createFormatter(formatterFactoryInput);
    }

    protected byte[] convertXslxToCsv(byte[] content) {
        try {
            ByteArrayInputStream inputStream = new ByteArrayInputStream(content);
            SpreadsheetMLPackage spreadsheet = (SpreadsheetMLPackage) SpreadsheetMLPackage.load(inputStream);
            WorksheetPart worksheetPart = spreadsheet.getWorkbookPart().getWorksheet(0);
            Worksheet ws = worksheetPart.getJaxbElement();
            SheetData sheetData = ws.getSheetData();

            List<String> csvRows = new ArrayList<>();
            for (Row r : sheetData.getRow()) {
                StringBuilder csvRow = new StringBuilder();
                for (Cell cell : r.getC()) {
                    String formatCellValue = null;
                    formatCellValue = cell.getV();
                    csvRow.append(formatCellValue).append(DEFAULT_SEPARATOR);
                }
                csvRows.add(csvRow.toString());
            }
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            IOUtils.writeLines(csvRows, "\n", outputStream, StandardCharsets.UTF_8);
            return outputStream.toByteArray();
        } catch (Docx4JException | Xlsx4jException | IOException e) {
            throw new ReportingException("Error while converting to CSV", e);
        }
    }

    protected class ReportTemplateWrapper implements ReportTemplate {
        ReportTemplate reportTemplate;

        public ReportTemplateWrapper(ReportTemplate reportTemplate) {
            this.reportTemplate = reportTemplate;
        }

        @Override
        public String getCode() {
            return reportTemplate.getCode();
        }

        @Override
        public String getDocumentName() {
            return reportTemplate.getDocumentName();
        }

        @Override
        public String getDocumentPath() {
            return reportTemplate.getDocumentPath();
        }

        @Override
        public InputStream getDocumentContent() {
            return reportTemplate.getDocumentContent();
        }

        @Override
        public ReportOutputType getOutputType() {
            return ReportOutputType.xlsx;
        }

        @Override
        public String getOutputNamePattern() {
            return null;
        }

        @Override
        public boolean isCustom() {
            return false;
        }

        @Override
        public CustomReport getCustomReport() {
            return null;
        }
    }
}
