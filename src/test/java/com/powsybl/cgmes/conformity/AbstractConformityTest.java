/**
 * Copyright (c) 2020, RTE (http://www.rte-france.com)
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

package com.powsybl.cgmes.conformity;

import com.google.common.jimfs.Configuration;
import com.google.common.jimfs.Jimfs;
import org.junit.After;
import org.junit.Before;
import org.xmlunit.builder.DiffBuilder;
import org.xmlunit.builder.Input;
import org.xmlunit.diff.Diff;

import javax.xml.transform.Source;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.FileSystem;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.junit.Assert.assertFalse;

/**
 * @author Mathieu Bague <mathieu.bague@rte-france.com>
 */
public abstract class AbstractConformityTest {

    private static final String RESOURCE_FILE = "/TestConfigurations_packageCASv2.0.zip";

    protected FileSystem fileSystem;
    protected Path tmpDir;

    @Before
    public void setup() throws IOException {
        fileSystem = Jimfs.newFileSystem(Configuration.unix());
        tmpDir = fileSystem.getPath("/tmp");
        Files.createDirectories(tmpDir);

        copyResources();
    }

    private void copyResources() throws IOException {
        ZipFile zipFile = new ZipFile(new File(getClass().getResource(RESOURCE_FILE).getFile()));
        for (String resource : getRequiredResources()) {
            ZipEntry zipEntry = zipFile.getEntry(resource.substring(1));
            try (InputStream stream = zipFile.getInputStream(zipEntry)) {
                Path dest = fileSystem.getPath(resource);
                Files.createDirectories(dest.getParent());
                Files.copy(stream, dest);
            }
        }
    }

    @After
    public void tearDown() throws IOException {
        fileSystem.close();
    }

    protected abstract List<String> getRequiredResources();

    protected static void compareXml(InputStream expected, InputStream actual) {
        Source control = Input.fromStream(expected).build();
        Source test = Input.fromStream(actual).build();
        Diff myDiff = DiffBuilder.compare(control).withTest(test).ignoreWhitespace().ignoreComments().build();
        boolean hasDiff = myDiff.hasDifferences();
        if (hasDiff) {
            System.err.println(myDiff.toString());
        }
        assertFalse(hasDiff);
    }
}
