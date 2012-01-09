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

import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Creates Sql script from xls file.
 *
 * @author vko
 * @version $Revision$ $Date$
 */
public class XlsToSql {

    public static void main(String[] args) {
        if (args.length < 2) {
            Logger.getLogger(XlsToSql.class.getName()).log(Level.INFO, "XLS to SQL convertor. Copyright (c) 2011 (javadev75@gmail.com)\nUsage: XlsToSql source.xls output.sql");
            return;
        }
        String xlsFiles = args[0].trim();
        String outSql = args[1].trim();
        Logger.getLogger(XlsToSql.class.getName()).log(Level.INFO, "xls files:" + xlsFiles);
        Logger.getLogger(XlsToSql.class.getName()).log(Level.INFO, "out sql:" + outSql);
        if (xlsFiles.isEmpty()) {
            throw new IllegalArgumentException("xlsFiles is empty");
        }
        if (outSql.isEmpty()) {
            throw new IllegalArgumentException("outSql is empty");
        }

        String[] cmdArgs = xlsFiles.split(",");
        List<String> fileLocations = new ArrayList<String>();
        for (String file : cmdArgs) {
            fileLocations.add(file.trim());
        }
        new SqlGenerator(fileLocations, outSql).generate();

        Logger.getLogger(XlsToSql.class.getName()).log(Level.INFO, outSql + " generated.");
    }
}
