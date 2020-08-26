/**
 * Copyright (c) 2020, RTE (http://www.rte-france.com)
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */
package com.powsybl.cgmes.validation.util;

import com.powsybl.cgmes.conversion.CgmesConversionContextExtension;
import com.powsybl.cgmes.conversion.CgmesModelExtension;
import com.powsybl.cgmes.model.CgmesModel;
import com.powsybl.cgmes.model.CgmesModelFactory;
import com.powsybl.cgmes.model.triplestore.CgmesModelTripleStore;
import com.powsybl.commons.PowsyblException;
import com.powsybl.commons.datasource.FileDataSource;
import com.powsybl.commons.datasource.ReadOnlyDataSource;
import com.powsybl.iidm.network.Bus;
import com.powsybl.iidm.network.Network;
import com.powsybl.iidm.network.Terminal;
import com.powsybl.triplestore.api.PropertyBag;
import com.powsybl.triplestore.api.PropertyBags;
import com.powsybl.cgmes.validation.CgmesValidationTool;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;


/**
 * @author Jérémy LABOUS <jlabous at silicom.fr>
 */
public final class SvUtils {

    private static final String TOPOLOGICAL_NODE = "TopologicalNode";
    private static final String DEPENDING_MODEL = "    <md:Model.DependentOn rdf:resource=\"%s\"/>%n";
    private static final String FULL_MODEL = "FullModel";
    private static final String PROFILE = "profile";

    public static String extractSVFile(String extractDest, String zipPackage, String separator) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(zipPackage);
        BufferedInputStream bufferedInputStream = new BufferedInputStream(fileInputStream);
        try (ZipInputStream zin = new ZipInputStream(bufferedInputStream)) {
            ZipEntry ze = null;
            while ((ze = zin.getNextEntry()) != null) {
                if (ze.getName().contains("_SV_")) {
                    String svFile = extractDest + separator + ze.getName();
                    try (OutputStream out = new FileOutputStream(svFile)) {
                        byte[] buffer = new byte[9000];
                        int len;
                        while ((len = zin.read(buffer)) != -1) {
                            out.write(buffer, 0, len);
                        }
                        return svFile;
                    }
                }
            }
        }
        throw new PowsyblException("SV File not found");
    }

    public static void compareSV(Network network, String baseName, String provider, SXSSFWorkbook wb) throws IOException {
        Map<String, Quintuplet> results = new HashMap<>();

        CgmesModelExtension cgmesModelExtension = network.getExtension(CgmesModelExtension.class);
        CgmesConversionContextExtension cgmesTerminalMappingExtension = network.getExtension(CgmesConversionContextExtension.class);
        for (PropertyBag p : cgmesModelExtension.getCgmesModel().topologicalNodes()) {
            Terminal t = cgmesTerminalMappingExtension.getContext().terminalMapping().findFromTopologicalNode(p.getId(TOPOLOGICAL_NODE));
            if (t != null) {
                Bus bus = t.getBusBreakerView().getConnectableBus();
                if (bus != null && StringUtils.isNotEmpty(bus.getVoltageLevel().getId())  && StringUtils.isNotEmpty(bus.getId())) {
                    results.computeIfAbsent(p.getId(TOPOLOGICAL_NODE), s -> {
                        Quintuplet quintuplet = new Quintuplet();
                        quintuplet.voltageLevelId = t.getVoltageLevel().getId();
                        quintuplet.busId = t.getBusBreakerView().getConnectableBus().getId();
                        quintuplet.equipmentId = t.getConnectable().getId();
                        quintuplet.connectedComponentNumber = bus.getConnectedComponent() != null ? bus.getConnectedComponent().getNum() : 0;
                        return quintuplet;
                    });
                }
            }
        }
        addFile(Paths.get(CgmesValidationTool.WORKING_DIR).resolve("datafiles"), baseName, results, (quintuplet, p) -> {
            quintuplet.v1 = p.asDouble(XlsUtils.V);
            quintuplet.angle1 = p.asDouble(XlsUtils.ANGLE);
        });
        addFile(Paths.get(CgmesValidationTool.WORKING_DIR).resolve("postImport_SSH"), baseName, results, (quintuplet, p) -> {
            quintuplet.v2 = p.asDouble(XlsUtils.V);
            quintuplet.angle2 = p.asDouble(XlsUtils.ANGLE);
        });
        addFile(Paths.get(CgmesValidationTool.WORKING_DIR).resolve("postLF_SSH"), baseName, results, (quintuplet, p) -> {
            quintuplet.v3 = p.asDouble(XlsUtils.V);
            quintuplet.angle3 = p.asDouble(XlsUtils.ANGLE);
            if (p.asDouble(XlsUtils.ANGLE) == 0) {
                XlsUtils.getDiffAngle()[0] = quintuplet.angle3 - quintuplet.angle2;
            }
        });
        addFile(Paths.get(CgmesValidationTool.WORKING_DIR).resolve("postImport_SV"), baseName, results, (quintuplet, p) -> {
            quintuplet.v4 = p.asDouble(XlsUtils.V);
            quintuplet.angle4 = p.asDouble(XlsUtils.ANGLE);
        });
        addFile(Paths.get(CgmesValidationTool.WORKING_DIR).resolve("postLF_SV"), baseName, results, (quintuplet, p) -> {
            quintuplet.v5 = p.asDouble(XlsUtils.V);
            quintuplet.angle5 = p.asDouble(XlsUtils.ANGLE);
            if (p.asDouble(XlsUtils.ANGLE) == 0) {
                XlsUtils.getDiffAngle()[1] = quintuplet.angle5 - quintuplet.angle4;
            }
        });
        XlsUtils.writeCompareSVResults(wb, results, provider);
        FileUtils.deleteDirectory(Paths.get(CgmesValidationTool.WORKING_DIR).resolve("datafiles").toFile());
        FileUtils.deleteDirectory(Paths.get(CgmesValidationTool.WORKING_DIR).resolve("postImport_SSH").toFile());
        FileUtils.deleteDirectory(Paths.get(CgmesValidationTool.WORKING_DIR).resolve("postLF_SSH").toFile());
        FileUtils.deleteDirectory(Paths.get(CgmesValidationTool.WORKING_DIR).resolve("postImport_SV").toFile());
        FileUtils.deleteDirectory(Paths.get(CgmesValidationTool.WORKING_DIR).resolve("postLF_SV").toFile());
    }

    private static void addFile(Path folder, String file, Map<String, Quintuplet> results, BiConsumer<Quintuplet, PropertyBag> consumer) {
        ReadOnlyDataSource ds = new FileDataSource(folder, file);
        CgmesModelTripleStore cgmes = (CgmesModelTripleStore) CgmesModelFactory.create(ds, "rdf4j");
        PropertyBags bags = cgmes.namedQuery("voltages");
        for (PropertyBag p : bags) {
            consumer.accept(results.computeIfAbsent(p.getId(TOPOLOGICAL_NODE), s -> new Quintuplet()), p);
        }
    }

    public static class Quintuplet {
        String voltageLevelId = "";
        String busId = "";
        String equipmentId = "";
        int connectedComponentNumber = 0;
        double v1 = Double.NaN;
        double v2 = Double.NaN;
        double v3 = Double.NaN;
        double v4 = Double.NaN;
        double v5 = Double.NaN;
        double angle1 = Double.NaN;
        double angle2 = Double.NaN;
        double angle3 = Double.NaN;
        double angle4 = Double.NaN;
        double angle5 = Double.NaN;
    }

    public static void writeSvFile(Network network, String version, Writer writer) throws IOException {
        CgmesModelExtension cgmesModelExtension = network.getExtension(CgmesModelExtension.class);
        CgmesConversionContextExtension cgmesTerminalMappingExtension = network.getExtension(CgmesConversionContextExtension.class);

        if (cgmesModelExtension != null) {
            writerHeader(writer);
            writeFullModel(cgmesModelExtension, version, writer);
            writeAngleTension(cgmesModelExtension, cgmesTerminalMappingExtension, writer);
            writer.write("</rdf:RDF>");
        }
    }

    private static void writerHeader(Writer writer) throws IOException {
        writer.write("\uFEFF<?xml version=\"1.0\" encoding=\"utf-8\"?>\n");
        writer.write("<rdf:RDF xmlns:cim=\"http://iec.ch/TC57/2013/CIM-schema-cim16#\"");
        writer.write(" xmlns:md=\"http://iec.ch/TC57/61970-552/ModelDescription/1#\" xmlns:entsoe=\"http://entsoe.eu/CIM/SchemaExtension/3/1#\"");
        writer.write(" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">\n");
    }

    private static void writeFullModel(CgmesModelExtension cgmesModelExtension, String version, Writer writer) throws IOException {
        CgmesModel cgmes = cgmesModelExtension.getCgmesModel();
        PropertyBags properties = cgmes.modelProfiles();

        String[] svProfile = new String[1];

        for (PropertyBag p : properties) {
            String tmp = p.get(PROFILE);
            if (tmp != null && tmp.contains("/StateVariables/")) {
                svProfile[0] = p.getId(FULL_MODEL);
            }
        }
        properties.removeIf(p -> p.getId(FULL_MODEL).equals(svProfile[0]));
        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'");
        Date date = new Date();

        writer.write("  <md:FullModel rdf:about=\"" + svProfile[0] + "\">\n");
        writer.write("    <md:Model.created>" + dateFormat.format(date) + "</md:Model.created>\n");
        writer.write("    <md:Model.createdBy>powsybl</md:Model.createdBy>\n");
        writer.write("    <md:Model.scenarioTime>" + dateFormat.format(date) + "</md:Model.scenarioTime>\n");
        writer.write("    <md:Model.description>Generated by PowSyBl for tests</md:Model.description>\n");
        writer.write("    <md:Model.modelingAuthoritySet>powsybl</md:Model.modelingAuthoritySet>\n");
        writer.write("    <md:Model.profile>http://entsoe.eu/CIM/StateVariables/4/1</md:Model.profile>\n");
        writer.write("    <md:Model.version>" + version + "</md:Model.version>\n");
        writer.write(String.format(DEPENDING_MODEL, properties.stream().filter(p -> p.get(PROFILE) != null && p.get(PROFILE).contains("/EquipmentCore/")).findFirst().map(p -> p.getId(FULL_MODEL)).orElse("MISSING_EQ")));
        writer.write(String.format(DEPENDING_MODEL, properties.stream().filter(p -> p.get(PROFILE) != null && p.get(PROFILE).contains("/Topology/")).findFirst().map(p -> p.getId(FULL_MODEL)).orElse("MISSING_TP")));
        writer.write(String.format(DEPENDING_MODEL, properties.stream().filter(p -> p.get(PROFILE) != null && p.get(PROFILE).contains("/SteadyStateHypothesis/")).findFirst().map(p -> p.getId(FULL_MODEL)).orElse("MISSING_SSH")));
        writer.write("  </md:FullModel>\n");
    }

    private static void writeAngleTension(CgmesModelExtension cgmesModelExtension, CgmesConversionContextExtension cgmesTerminalMappingExtension, Writer writer) throws IOException {
        int counter = 0;
        for (PropertyBag p : cgmesModelExtension.getCgmesModel().topologicalNodes()) {
            Terminal t = cgmesTerminalMappingExtension.getContext().terminalMapping().findFromTopologicalNode(p.getId("TopologicalNode"));
            if (t != null) {
                Bus bus = t.getBusBreakerView().getConnectableBus();
                if (bus != null) {
                    writer.write(String.format("  <cim:SvVoltage rdf:ID=\"%d\">%n", counter));
                    writer.write(String.format("    <cim:SvVoltage.angle>%s</cim:SvVoltage.angle>%n", NumberFormat.getInstance(Locale.US).format(bus.getAngle())));
                    writer.write(String.format("    <cim:SvVoltage.v>%s</cim:SvVoltage.v>%n", NumberFormat.getInstance(Locale.US).format(bus.getV())));
                    writer.write(String.format("    <cim:SvVoltage.TopologicalNode rdf:resource=\"#%s\" />%n", p.getId("TopologicalNode")));
                    writer.write("  </cim:SvVoltage>\n");
                    counter++;
                }
            }
        }
    }

    private SvUtils() {
    }
}
