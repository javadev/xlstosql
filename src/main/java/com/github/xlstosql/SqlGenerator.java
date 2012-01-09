/*
 * $Id$
 *
 * Copyright 2012 Valentyn Kolesnikov
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.github.xlstosql;

import java.util.ArrayList;
import java.util.List;

import java.io.File; 
import java.io.FileOutputStream; 
import java.util.Date; 
import java.util.Calendar;
import java.text.SimpleDateFormat;
import jxl.*;
import org.apache.commons.io.IOUtils;

import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * SqlGenerator.
 *
 * @author vko
 * @version $Revision$ $Date$
 */
public class SqlGenerator {

    private final List<String> files;
    private final String outSql;
    private StringBuilder result;
    private String sheetName;
    private Cell[] headers;
    private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    /** Instance logger */
    private Logger log;

    public SqlGenerator(List<String> files, String outSql) {
        this.files = files;
        this.outSql = outSql;
    }

    /**
     * By default, return a <code>SystemStreamLog</code> logger.
     *
     * @see org.apache.maven.plugin.Mojo#getLog()
     */
    public Logger getLog() {
        if (log == null) {
            log = Logger.getLogger(SqlGenerator.class.getName());
        }
        return log;
    }

    public void generate() {
        readTables();
    }

    private void readTables() {
        result = new StringBuilder();
        try {
        for (String fileName : files) {
            Workbook workbook = Workbook.getWorkbook(new File(fileName));
            List<String> sheetNames = new ArrayList<String>();
            for (Sheet sheet : workbook.getSheets()) {
                sheetNames.add(sheet.getName());
                headers = sheet.getRow(0);
                for (int indexRow = 1; indexRow < sheet.getRows(); indexRow += 1) {
                    sheetName = sheet.getName();
                    appendHeaders();
                    Cell[] datas = sheet.getRow(indexRow);
                    if (datas != null && datas.length > 0) {
                        result.append(translateValue(0, cellToString(datas[0])));
                    }
                    for (int indexCell = 1; indexCell < headers.length; indexCell += 1) {
                        if (isValidDictHeader(headers[indexCell].getContents())) {
                            result.append(", " + translateValue(indexCell, cellToString(datas[indexCell])));
                        }
                    }
                    if (isDict()) {
                        Calendar today = Calendar.getInstance();
                        Calendar from = (Calendar) today.clone();
                        Calendar to = (Calendar) today.clone();
                        from.add(Calendar.DATE, -1000);
                        to.add(Calendar.DATE, 1000);
                        result.append(", '" + sheetName.substring(5)
                            + "', 1, " + toSqlFormat(from) + ", " + toSqlFormat(to) + ", " + toSqlFormat(today));
                    }
                    result.append(")\n");
                }
            }
            workbook.close();
            IOUtils.write(result.toString(), new FileOutputStream(outSql), "cp1251") ;
    getLog().log(Level.INFO, "read " + fileName.replaceFirst(".*/(.*)", "$1") + " - " + sheetNames);
        }
        } catch (Exception ex) {
            getLog().log(Level.WARNING, ex.getMessage(), ex);
        }
    }

    private String toSqlFormat(Calendar date) {
        return "'" + sdf.format(date.getTime()) + ".0'";
    }

    private void appendHeaders() {
        result.append("insert into " + getTableName() + " (");
        if (headers != null && headers.length > 0) {
            result.append(translateHeader(0));
        }
        for (int indexCell = 1; indexCell < headers.length; indexCell += 1) {
            if (isValidDictHeader(headers[indexCell].getContents())) {
                result.append(", " + translateHeader(indexCell));
            }
        }
        if (isDict()) {
            result.append(", name, valid, fromdate, todate, lastmodificationdate");
        }
        result.append(") values (");
    }

    private String translateHeader(int indexHeader) {
        if ("russian".equals(headers[indexHeader].getContents())) {
            return "dvalue, language";
        }
        return headers[indexHeader].getContents().replaceFirst("^key$", "dkey")
            .replaceFirst("^value$", "dvalue");
    }
    private String translateValue(int indexHeader, String value) {
        if ("russian".equals(headers[indexHeader].getContents())) {
            return value + ", 'ru'";
        }
        return value;
    }
    private boolean isValidDictHeader(String headerItem) {
        if (!isDict()) {
            return true;
        }
        return headerItem.matches("(key|value|expkey|expkey2|expkey3|russian|ukrainian)");
    }

    private boolean isDict() {
        return sheetName.startsWith("dict_");
    }

    private String getTableName() {
        return isDict() ? "dictionary_data" : sheetName;
    }

    private String cellToString(Cell cell) {
        if (cell.getType().equals(CellType.NUMBER)
            || cell.getType().equals(CellType.NUMBER_FORMULA)) {
            return cell.getContents().replaceFirst(",", ".").replaceFirst("%", "");
        }
        return "'" + cell.getContents().replaceAll("'", "''") + "'";
    }
}
