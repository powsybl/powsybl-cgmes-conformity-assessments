/**
 * Copyright (c) 2020, RTE (http://www.rte-france.com)
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */
package com.powsybl.cgmes.validation.util;

import com.powsybl.cgmes.conversion.CgmesImport;
import com.powsybl.commons.PowsyblException;
import com.powsybl.commons.io.table.*;
import com.powsybl.iidm.import_.ImportConfig;
import com.powsybl.iidm.import_.Importers;
import com.powsybl.iidm.network.Network;
import com.powsybl.loadflow.LoadFlow;
import com.powsybl.loadflow.LoadFlowParameters;
import com.powsybl.loadflow.LoadFlowResult;
import com.powsybl.loadflow.json.JsonLoadFlowParameters;
import com.powsybl.openloadflow.OpenLoadFlowParameters;
import com.powsybl.tools.ToolRunningContext;
import com.powsybl.cgmes.validation.CgmesValidationTool;
import org.apache.commons.cli.CommandLine;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UncheckedIOException;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Properties;

/**
 * @author Jérémy LABOUS <jlabous at silicom.fr>
 */
public final class CgmesUtils {

    public static Network importFile(boolean skipPostProc, Path caseFile, String baseName,
                                     boolean writeFile, Properties inputParams, ToolRunningContext context) throws IOException {
        ImportConfig importConfig = (!skipPostProc) ? ImportConfig.load() : new ImportConfig();

        if (writeFile) {
            context.getOutputStream().println("Loading network '" + caseFile + "'");
        }
        Network network = Importers.loadNetwork(caseFile, context.getShortTimeExecutionComputationManager(), importConfig, inputParams);
        if (network == null) {
            throw new PowsyblException("import fail");
        }

        if (writeFile) {
            Path dir = Paths.get("workingDir" + context.getFileSystem().getSeparator()
                    + "postImport_" + inputParams.getProperty(CgmesImport.PROFILE_USED_FOR_INITIAL_STATE_VALUES));
            Files.createDirectories(dir);
            try (Writer writer = Files.newBufferedWriter(dir.resolve(baseName), StandardCharsets.UTF_8)) {
                SvUtils.writeSvFile(network, "002", writer);
            }
        }
        return network;
    }

    public static LoadFlowResult runLoadFlow(Network network, String baseName, Properties inputParams, CommandLine line, ToolRunningContext context) throws IOException {
        LoadFlowParameters params = LoadFlowParameters.load();
        OpenLoadFlowParameters openLoadFlowParameters = new OpenLoadFlowParameters()
                .setVoltageRemoteControl(true)
                .setThrowsExceptionInCaseOfSlackDistributionFailure(false);
        params.addExtension(OpenLoadFlowParameters.class, openLoadFlowParameters);
        if (line.hasOption(CgmesValidationTool.PARAMETERS_FILE)) {
            Path parametersFile = context.getFileSystem().getPath(line.getOptionValue(CgmesValidationTool.PARAMETERS_FILE));
            JsonLoadFlowParameters.update(params, parametersFile);
        }

        LoadFlowResult lfResult = LoadFlow.find("OpenLoadFlow").run(network, params);
        printResult(lfResult, context);

        Path dir = Paths.get("workingDir" + context.getFileSystem().getSeparator()
                + "postLF_" + inputParams.getProperty(CgmesImport.PROFILE_USED_FOR_INITIAL_STATE_VALUES));
        Files.createDirectories(dir);
        try (Writer writer = Files.newBufferedWriter(dir.resolve(baseName), StandardCharsets.UTF_8)) {
            SvUtils.writeSvFile(network, "002", writer);
        }
        return lfResult;
    }

    private static void printResult(LoadFlowResult result, ToolRunningContext context) {
        Writer writer = new OutputStreamWriter(context.getOutputStream());
        AsciiTableFormatterFactory asciiTableFormatterFactory = new AsciiTableFormatterFactory();
        printLoadFlowResult(result, writer, asciiTableFormatterFactory, TableFormatterConfig.load());
    }

    private static void printLoadFlowResult(LoadFlowResult result, Writer writer, TableFormatterFactory formatterFactory,
                                     TableFormatterConfig formatterConfig) {
        try (TableFormatter formatter = formatterFactory.create(writer,
                "loadflow results",
                formatterConfig,
                new Column("Result"),
                new Column("Metrics"))) {
            formatter.writeCell(result.isOk());
            formatter.writeCell(result.getMetrics().toString());
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    private CgmesUtils() {
    }
}
