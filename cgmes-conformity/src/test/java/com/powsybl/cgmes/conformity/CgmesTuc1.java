/**
 * Copyright (c) 2020, RTE (http://www.rte-france.com)
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

package com.powsybl.cgmes.conformity;

import com.powsybl.cgmes.model.CgmesModel;
import com.powsybl.cgmes.model.CgmesModelFactory;
import com.powsybl.cgmes.model.CgmesSubset;
import com.powsybl.commons.datasource.FileDataSource;
import com.powsybl.commons.datasource.ZipFileDataSource;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collections;
import java.util.List;

/**
 * @author Mathieu Bague <mathieu.bague@rte-france.com>
 */
public class CgmesTuc1 extends AbstractConformityTest {

    private ZipFileDataSource dataSource;

    private CgmesModel model;

    @Override
    public void setup() throws IOException {
        super.setup();

        dataSource = new ZipFileDataSource(fileSystem.getPath("/FullGrid/CGMES_v2.4.15_FullGridTestConfiguration_BD_v1.zip"));
        model = CgmesModelFactory.create(dataSource);
    }

    @Override
    protected List<String> getRequiredResources() {
        return Collections.singletonList("/FullGrid/CGMES_v2.4.15_FullGridTestConfiguration_BD_v1.zip");
    }

    @Test
    public void testEQ() throws IOException {
        test("20171002T0930Z_ENTSO-E_EQ_BD_1.xml", CgmesSubset.EQUIPMENT_BOUNDARY);
    }

    @Test
    public void testTP() throws IOException {
        test("20171002T0930Z_ENTSO-E_TP_BD_1.xml", CgmesSubset.TOPOLOGY_BOUNDARY);
    }

    private void test(String filename, CgmesSubset subset) throws IOException {
        Path outputPath = tmpDir.resolve(filename);
        model.write(new FileDataSource(tmpDir, ""), subset);

        try (InputStream expectedStream = dataSource.newInputStream(filename);
             InputStream actualStream = Files.newInputStream(outputPath)) {
            compareXml(expectedStream, actualStream);
        }
    }
}
