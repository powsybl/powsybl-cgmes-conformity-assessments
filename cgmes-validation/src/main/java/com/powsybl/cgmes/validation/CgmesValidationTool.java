/**
 * Copyright (c) 2020, RTE (http://www.rte-france.com)
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */
package com.powsybl.cgmes.validation;

import com.google.auto.service.AutoService;
import com.powsybl.cgmes.conversion.CgmesImport;
import com.powsybl.cgmes.model.CgmesModelException;
import com.powsybl.iidm.network.Network;
import com.powsybl.iidm.tools.ConversionToolUtils;
import com.powsybl.loadflow.LoadFlowResult;
import com.powsybl.tools.Command;
import com.powsybl.tools.Tool;
import com.powsybl.tools.ToolRunningContext;
import com.powsybl.cgmes.validation.util.SvUtils;
import com.powsybl.cgmes.validation.util.CgmesUtils;
import com.powsybl.cgmes.validation.util.XlsUtils;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.*;
import java.util.regex.Pattern;

import static com.powsybl.iidm.tools.ConversionToolUtils.readProperties;

/**
 * @author Jérémy LABOUS <jlabous at silicom.fr>
 */
@AutoService(Tool.class)
public class CgmesValidationTool implements Tool {

    public static final String PARAMETERS_FILE = "parameters-file";
    public static final String SKIP_POSTPROC = "skip-postproc";
    public static final String DATASET_PATH = "dataset-path";
    public static final String EXPLOITED_HOUR = "exploited-hour";
    public static final String RESULT_PATH = "result-path";
    public static final String RESULT_FILENAME = "result-filename";

    public static final String WORKING_DIR = "workingDir";
    public static final String PROVIDER_KEY = "provider";
    public static final String TOPOLOGY_KEY = "topology";
    public static final String ERROR_KEY = "error";
    public static final String CASEFILE_KEY = "caseFile";
    public static final String LOADFLOW_KEY = "lfResult";

    private Properties inputParams;

    @Override
    public Command getCommand() {
        return new Command() {
            @Override
            public String getName() {
                return "cgmes-validation";
            }

            @Override
            public String getTheme() {
                return "CGMES validation";
            }

            @Override
            public String getDescription() {
                return "Import CGMES, generate SV files, run loadflow and compare SV files";
            }

            @Override
            public Options getOptions() {
                Options options = new Options();
                options.addOption(Option.builder().longOpt(PARAMETERS_FILE)
                        .desc("loadflow parameters as JSON file")
                        .hasArg()
                        .argName("FILE")
                        .build());
                options.addOption(Option.builder().longOpt(SKIP_POSTPROC)
                        .desc("skip network importer post processors (when configured)")
                        .build());
                options.addOption(Option.builder().longOpt(DATASET_PATH)
                        .desc("configure the path of the dataset")
                        .hasArg()
                        .argName("DATASET")
                        .required()
                        .build());
                options.addOption(Option.builder().longOpt(EXPLOITED_HOUR)
                        .desc("configure the exploited hour for the validation")
                        .hasArg()
                        .argName("HOUR")
                        .required()
                        .build());
                options.addOption(Option.builder().longOpt(RESULT_PATH)
                        .desc("configure the path for result file")
                        .hasArg()
                        .argName("RESULT")
                        .required()
                        .build());
                options.addOption(Option.builder().longOpt(RESULT_FILENAME)
                        .desc("configure the result file name")
                        .hasArg()
                        .argName("FILENAME")
                        .build());
                options.addOption(ConversionToolUtils.createImportParametersFileOption());
                options.addOption(ConversionToolUtils.createImportParameterOption());
                options.addOption(ConversionToolUtils.createExportParametersFileOption());
                options.addOption(ConversionToolUtils.createExportParameterOption());
                return options;
            }

            @Override
            public String getUsageFooter() {
                return null;
            }
        };
    }

    @Override
    public void run(CommandLine line, ToolRunningContext context) throws IOException {
        Map<String, Map<String, String>> results = new HashMap<>();
        Path dataSetFolder = context.getFileSystem().getPath(line.getOptionValue(DATASET_PATH));
        String resultFolder = line.getOptionValue(RESULT_PATH);
        String exploitedHour = "_" + line.getOptionValue(EXPLOITED_HOUR) + "_";
        String resultFileName = Optional.ofNullable(line.getOptionValue(RESULT_FILENAME)).orElse("Results")  + ".xlsx";

        inputParams = readProperties(line, ConversionToolUtils.OptionType.IMPORT, context);

        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            wb.createSheet("Results");
            for (String folder : Objects.requireNonNull(dataSetFolder.toFile().list())) {
                Map<String, String> result = new HashMap<>();
                result.put(CASEFILE_KEY, folder);
                String[] splitFolder = folder.split("_");
                if (splitFolder.length <= 5 || Pattern.matches("[0-9]*", splitFolder[5])) {
                    result.put(PROVIDER_KEY, splitFolder[3]);
                } else {
                    result.put(PROVIDER_KEY, splitFolder[3] + "_" + splitFolder[4]);
                }
                String dataFiles = WORKING_DIR + context.getFileSystem().getSeparator() + "datafiles";
                if (new File(dataFiles).mkdirs()) {
                    Path dataFolder = context.getFileSystem().getPath(dataSetFolder.toString() + context.getFileSystem().getSeparator() + folder);
                    Path caseFile = null;
                    for (String zipPackage : Objects.requireNonNull(dataFolder.toFile().list())) {
                        if (zipPackage.contains(".zip") && zipPackage.contains(exploitedHour)) {
                            caseFile = context.getFileSystem().getPath(dataFolder + context.getFileSystem().getSeparator() + zipPackage);
                        }
                    }
                    processFile(dataFiles, result.get(PROVIDER_KEY), caseFile, result, wb, line, context);
                    results.put(folder, result);
                }
            }

            XlsUtils.generateResults(wb, results);
            try (FileOutputStream fos = new FileOutputStream(resultFileName)) {
                wb.write(fos);
            }
        }
        File resultFile = new File(resultFileName);
        FileUtils.copyFile(resultFile, new File(resultFolder + context.getFileSystem().getSeparator() + resultFileName));
        FileUtils.deleteQuietly(resultFile);
        FileUtils.deleteDirectory(new File(WORKING_DIR));
        context.getOutputStream().println(results);
    }

    private void processFile(String dataFiles, String provider, Path caseFile, Map<String, String> result,
                             SXSSFWorkbook wb, CommandLine line, ToolRunningContext context) throws IOException {
        boolean skipPostProc = line.hasOption(SKIP_POSTPROC);
        Path svFile = null;
        if (caseFile != null) {
            try {
                svFile = context.getFileSystem().getPath(SvUtils.extractSVFile(dataFiles, caseFile.toString(), context.getFileSystem().getSeparator()));
            } catch (CgmesModelException cme) {
                result.put(ERROR_KEY, cme.getMessage());
            }
        } else {
            result.put(ERROR_KEY, "SV file not found for the provider " + provider);
            return;
        }

        inputParams.setProperty(CgmesImport.PROFILE_USED_FOR_INITIAL_STATE_VALUES, "SSH");

        Network network;
        Network networkPI;
        String filename;
        try {
            Objects.requireNonNull(svFile, "SV file is null");
            filename = svFile.getFileName().toString();
            network = CgmesUtils.importFile(skipPostProc, caseFile, filename, true, inputParams, context);
            networkPI = CgmesUtils.importFile(skipPostProc, caseFile, filename, false, inputParams, context);

            network.getVoltageLevelStream().findFirst().ifPresent(vl -> result.put(TOPOLOGY_KEY, vl.getTopologyKind().toString()));
        } catch (Exception e) {
            result.put(ERROR_KEY, "error during cgmes import SSH");
            return;
        }

        try {
            LoadFlowResult lfResult = CgmesUtils.runLoadFlow(network, filename, inputParams, line, context);
            result.put(LOADFLOW_KEY, lfResult.getMetrics().get("network_0_status"));
        } catch (Exception e) {
            result.put(ERROR_KEY, "error during loadflow SSH");
        }

        inputParams.setProperty(CgmesImport.PROFILE_USED_FOR_INITIAL_STATE_VALUES, "SV");
        try {
            network = CgmesUtils.importFile(skipPostProc, caseFile, filename, true, inputParams, context);
        } catch (Exception e) {
            result.put(ERROR_KEY, "error during cgmes import SV");
            return;
        }

        try {
            LoadFlowResult lfResult = CgmesUtils.runLoadFlow(network, filename, inputParams, line, context);
            result.put(LOADFLOW_KEY, lfResult.getMetrics().get("network_0_status"));
        } catch (Exception e) {
            result.put(ERROR_KEY, "error during loadflow SV");
        }

        SvUtils.compareSV(networkPI, filename, provider, wb);
    }
}
