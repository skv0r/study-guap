#!/usr/bin/env python3
"""Create a KNIME 5 linear-regression workflow on disk (courier delivery)."""
from __future__ import annotations

import shutil
from pathlib import Path
from textwrap import dedent

LAB = Path(__file__).resolve().parents[1]
CSV_SRC = LAB / "data" / "courier_delivery.csv"
WF = LAB / "workflow" / "LR1_Courier_LinReg"
WS_WF = Path(r"C:\Users\skvor\knime-workspace") / "LR1_Courier_LinReg"

NS = 'xmlns="http://www.knime.org/2008/09/XMLConfig" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.knime.org/2008/09/XMLConfig http://www.knime.org/XMLConfig_2008_09.xsd"'
BUNDLE = """    <entry key="node-bundle-name" type="xstring" value="KNIME Base Nodes"/>
    <entry key="node-bundle-symbolic-name" type="xstring" value="org.knime.base"/>
    <entry key="node-bundle-vendor" type="xstring" value="KNIME AG, Zurich, Switzerland"/>
    <entry key="node-bundle-version" type="xstring" value="5.12.0"/>
    <entry key="node-feature-name" type="xstring" value="KNIME Base nodes"/>
    <entry key="node-feature-symbolic-name" type="xstring" value="org.knime.features.base.feature.group"/>
    <entry key="node-feature-vendor" type="xstring" value="KNIME AG, Zurich, Switzerland"/>
    <entry key="node-feature-version" type="xstring" value="5.12.0"/>"""


def header():
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<config {NS} key="settings.xml">
    <entry key="node_file" type="xstring" value="settings.xml"/>
    <config key="flow_stack"/>
    <config key="internal_node_subsettings">
        <entry key="memory_policy" type="xstring" value="CacheSmallInMemory"/>
    </config>'''


def footer(factory, name, extra=""):
    return f'''    <entry key="customDescription" type="xstring" isnull="true" value=""/>
    <entry key="state" type="xstring" value="CONFIGURED"/>
    <entry key="factory" type="xstring" value="{factory}"/>
    <entry key="node-name" type="xstring" value="{name}"/>
{BUNDLE}
    <config key="factory_settings"/>
    {extra}
    <entry key="name" type="xstring" value="{name}"/>
    <entry key="hasContent" type="xboolean" value="false"/>
    <entry key="isInactive" type="xboolean" value="false"/>
    <config key="ports"/>
    <config key="filestores">
        <entry key="file_store_location" type="xstring" isnull="true" value=""/>
        <entry key="file_store_id" type="xstring" isnull="true" value=""/>
    </config>
</config>
'''


def col_prod(idx, name, jtype):
    if jtype == "Integer":
        conv = "org.knime.core.data.def.IntCell$IntCellFactory.create(class java.lang.Integer)"
        dst = "Number (integer)"
        src = "java.lang.Integer"
        cname = "Integer"
    else:
        conv = "org.knime.core.data.def.DoubleCell$DoubleCellFactory.create(class java.lang.Double)"
        dst = "Number (double)"
        src = "java.lang.Double"
        cname = "Double"
    return f'''                    <config key="{idx}">
                        <entry key="name" type="xstring" value="{name}"/>
                        <entry key="has_type" type="xboolean" value="true"/>
                        <config key="type">
                            <entry key="class" type="xstring" value="{src}"/>
                        </config>
                    </config>'''


def col_trans(idx, name, jtype):
    if jtype == "Integer":
        conv = "org.knime.core.data.def.IntCell$IntCellFactory.create(class java.lang.Integer)"
        dst = "Number (integer)"
        src = "java.lang.Integer"
        cname = "Integer"
    else:
        conv = "org.knime.core.data.def.DoubleCell$DoubleCellFactory.create(class java.lang.Double)"
        dst = "Number (double)"
        src = "java.lang.Double"
        cname = "Double"
    return f'''                    <config key="{idx}">
                        <config key="external_spec">
                            <entry key="name" type="xstring" value="{name}"/>
                            <entry key="has_type" type="xboolean" value="true"/>
                            <config key="type">
                                <entry key="class" type="xstring" value="{src}"/>
                            </config>
                        </config>
                        <entry key="name" type="xstring" value="{name}"/>
                        <entry key="keep" type="xboolean" value="true"/>
                        <entry key="position" type="xint" value="{idx}"/>
                        <config key="production_path">
                            <entry key="_converter" type="xstring" value="{conv}"/>
                            <entry key="_converter_src" type="xstring" value="{src}"/>
                            <entry key="_converter_dst" type="xstring" value="{dst}"/>
                            <entry key="_converter_name" type="xstring" value="{cname}"/>
                            <config key="_converter_config"/>
                            <entry key="_producer" type="xstring" value="class {src}-&gt;{src}"/>
                            <entry key="_producer_src" type="xstring" value="{src}"/>
                            <entry key="_producer_dst" type="xstring" value="{src}"/>
                            <entry key="_producer_name" type="xstring" value="{src}"/>
                            <config key="_producer_config"/>
                        </config>
                    </config>'''


def csv_reader():
    cols = [("order_id", "Integer"), ("distance_km", "Double"), ("time_min", "Double")]
    specs = "\n".join(col_prod(i, n, t) for i, (n, t) in enumerate(cols))
    trans = "\n".join(col_trans(i, n, t) for i, (n, t) in enumerate(cols))
    path = "data/courier_delivery.csv"
    return f'''{header()}
    <config key="model">
        <config key="settings">
            <config key="file_selection_Internals">
                <entry key="SettingsModelID" type="xstring" value="SMID_ReaderFileChooser"/>
                <entry key="EnabledStatus" type="xboolean" value="true"/>
            </config>
            <config key="file_selection">
                <config key="file_system_chooser__Internals">
                    <entry key="has_fs_port" type="xboolean" value="false"/>
                    <entry key="overwritten_by_variable" type="xboolean" value="false"/>
                    <entry key="convenience_fs_category" type="xstring" value="RELATIVE"/>
                    <entry key="relative_to" type="xstring" value="knime.workflow"/>
                    <entry key="mountpoint" type="xstring" value="LOCAL"/>
                    <entry key="spaceId" type="xstring" value=""/>
                    <entry key="spaceName" type="xstring" value=""/>
                    <entry key="custom_url_timeout" type="xint" value="1000"/>
                    <entry key="connected_fs" type="xboolean" value="true"/>
                </config>
                <config key="path">
                    <entry key="location_present" type="xboolean" value="true"/>
                    <entry key="file_system_type" type="xstring" value="RELATIVE"/>
                    <entry key="file_system_specifier" type="xstring" value="knime.workflow"/>
                    <entry key="path" type="xstring" value="{path}"/>
                </config>
                <config key="filter_mode_Internals">
                    <entry key="SettingsModelID" type="xstring" value="SMID_FilterMode"/>
                    <entry key="EnabledStatus" type="xboolean" value="true"/>
                </config>
                <config key="filter_mode">
                    <entry key="filter_mode" type="xstring" value="FILE"/>
                    <entry key="include_subfolders" type="xboolean" value="false"/>
                    <config key="filter_options">
                        <entry key="filter_files_extension" type="xboolean" value="false"/>
                        <entry key="files_extension_expression" type="xstring" value=""/>
                        <entry key="files_extension_case_sensitive" type="xboolean" value="false"/>
                        <entry key="filter_files_name" type="xboolean" value="false"/>
                        <entry key="files_name_expression" type="xstring" value="*"/>
                        <entry key="files_name_case_sensitive" type="xboolean" value="false"/>
                        <entry key="files_name_filter_type" type="xstring" value="WILDCARD"/>
                        <entry key="include_hidden_files" type="xboolean" value="false"/>
                        <entry key="include_special_files" type="xboolean" value="true"/>
                        <entry key="filter_folders_name" type="xboolean" value="false"/>
                        <entry key="folders_name_expression" type="xstring" value="*"/>
                        <entry key="folders_name_case_sensitive" type="xboolean" value="false"/>
                        <entry key="folders_name_filter_type" type="xstring" value="WILDCARD"/>
                        <entry key="include_hidden_folders" type="xboolean" value="false"/>
                        <entry key="follow_links" type="xboolean" value="true"/>
                    </config>
                </config>
            </config>
            <entry key="has_column_header" type="xboolean" value="true"/>
            <entry key="has_row_id" type="xboolean" value="false"/>
            <entry key="support_short_data_rows" type="xboolean" value="false"/>
            <entry key="skip_empty_data_rows" type="xboolean" value="false"/>
            <entry key="prepend_file_idx_to_row_id" type="xboolean" value="false"/>
            <entry key="comment_char" type="xstring" value="#"/>
            <entry key="column_delimiter" type="xstring" value=","/>
            <entry key="quote_char" type="xstring" value="&quot;"/>
            <entry key="quote_escape_char" type="xstring" value="&quot;"/>
            <entry key="use_line_break_row_delimiter" type="xboolean" value="true"/>
            <entry key="row_delimiter" type="xstring" value="%%00013%%00010"/>
            <entry key="autodetect_buffer_size" type="xint" value="1048576"/>
        </config>
        <config key="table_spec_config_Internals">
            <entry key="version" type="xstring" value="V4_4"/>
            <config key="individual_specs">
                <config key="{path}">
                    <entry key="num_columns" type="xint" value="3"/>
{specs}
                </config>
            </config>
            <config key="table_transformations">
                <config key="{path}">
                    <entry key="num_columns" type="xint" value="3"/>
                    <entry key="skip_empty_columns" type="xboolean" value="false"/>
                    <config key="columns">
{trans}
                    </config>
                </config>
            </config>
        </config>
    </config>
''' + footer(
        "org.knime.base.node.io.filehandling.csv.reader.CSVTableReaderNodeFactory",
        "CSV Reader",
        extra='<config key="node_creation_config"><config key="File System Connection"/><config key="Data Table"/></config>',
    )


def color_manager():
    return f'''{header()}
    <config key="model">
        <entry key="selected_column" type="xstring" value="distance_km"/>
        <entry key="min_color" type="xint" value="16776960"/>
        <entry key="max_color" type="xint" value="16711680"/>
    </config>
{footer("org.knime.base.node.viz.property.color.ColorManager2NodeFactory", "Color Manager")}
'''


def partitioning():
    return f'''{header()}
    <config key="model">
        <entry key="method" type="xstring" value="Relative"/>
        <entry key="samplingMethod" type="xstring" value="Linear"/>
        <entry key="fraction" type="xdouble" value="0.2"/>
        <entry key="count" type="xint" value="20"/>
        <entry key="random_seed" type="xstring" isnull="true" value=""/>
        <entry key="class_column" type="xstring" isnull="true" value=""/>
    </config>
{footer("org.knime.base.node.preproc.partition.PartitionNodeFactory", "Partitioning")}
'''


def linreg():
    return f'''{header()}
    <config key="model">
        <entry key="target" type="xstring" value="time_min"/>
        <entry key="include_constant" type="xboolean" value="true"/>
        <entry key="offset_value" type="xdouble" value="0.0"/>
        <entry key="missing_value_handling" type="xstring" value="fail"/>
        <config key="column_filter">
            <entry key="filter-type" type="xstring" value="STANDARD"/>
            <config key="included_names">
                <entry key="array-size" type="xint" value="1"/>
                <entry key="0" type="xstring" value="distance_km"/>
            </config>
            <config key="excluded_names">
                <entry key="array-size" type="xint" value="2"/>
                <entry key="0" type="xstring" value="order_id"/>
                <entry key="1" type="xstring" value="time_min"/>
            </config>
            <entry key="enforce_option" type="xstring" value="EnforceInclusion"/>
            <config key="name_pattern">
                <entry key="pattern" type="xstring" value=""/>
                <entry key="type" type="xstring" value="Wildcard"/>
                <entry key="caseSensitive" type="xboolean" value="true"/>
                <entry key="excludeMatching" type="xboolean" value="false"/>
            </config>
            <config key="datatype">
                <config key="typelist"/>
            </config>
        </config>
    </config>
{footer("org.knime.base.node.mine.regression.linear2.learner.LinReg2LearnerNodeFactory2", "Linear Regression Learner")}
'''


def predictor():
    return f'''{header()}
    <config key="model">
        <entry key="has_custom_prediction_name" type="xboolean" value="false"/>
        <entry key="custom_prediction_name" type="xstring" value="Prediction (time_min)"/>
        <entry key="include_probabilites" type="xboolean" value="false"/>
        <entry key="propability_columns_suffix" type="xstring" value=""/>
    </config>
{footer("org.knime.base.node.mine.regression.predict3.RegressionPredictorNodeFactory2", "Regression Predictor")}
'''


def scorer():
    return f'''{header()}
    <config key="model">
        <entry key="reference" type="xstring" value="time_min"/>
        <entry key="predicted" type="xstring" value="Prediction (time_min)"/>
        <entry key="output column" type="xstring" value=""/>
        <entry key="number_of_predictors" type="xint" value="1"/>
    </config>
{footer("org.knime.base.node.mine.scorer.numeric2.NumericScorer2NodeFactory", "Numeric Scorer")}
'''


def scatter():
    return f'''{header()}
    <config key="model">
        <entry key="generateImage" type="xboolean" value="true"/>
        <entry key="width" type="xint" value="800"/>
        <entry key="height" type="xint" value="600"/>
    </config>
    <config key="view">
        <entry key="allowImageDownload" type="xboolean" value="true"/>
        <entry key="axisExtentMethod" type="xstring" value="AUTO"/>
        <entry key="dataPointSize" type="xint" value="8"/>
        <entry key="enableAnimation" type="xboolean" value="false"/>
        <entry key="enableDataZoom" type="xboolean" value="true"/>
        <entry key="maxRows" type="xint" value="12500"/>
        <entry key="publishSelection" type="xboolean" value="true"/>
        <config key="referenceLines">
            <entry key="null_Internals" type="xboolean" value="true"/>
        </config>
        <entry key="showTooltip" type="xboolean" value="true"/>
        <entry key="subscribeToSelection" type="xboolean" value="true"/>
        <entry key="title" type="xstring" value="Scatter Plot"/>
        <entry key="xAxisColumn" type="xstring" value="distance_km"/>
        <entry key="xAxisLabel" type="xstring" value="distance_km"/>
        <entry key="xAxisScale" type="xstring" value="VALUE"/>
        <entry key="yAxisColumn" type="xstring" value="time_min"/>
        <entry key="yAxisLabel" type="xstring" value="time_min"/>
        <entry key="yAxisScale" type="xstring" value="VALUE"/>
    </config>
    <entry key="customDescription" type="xstring" isnull="true" value=""/>
    <entry key="state" type="xstring" value="CONFIGURED"/>
    <entry key="factory" type="xstring" value="org.knime.base.views.node.scatterplot.ScatterPlotNodeFactory"/>
    <entry key="node-name" type="xstring" value="Scatter Plot"/>
    <entry key="node-bundle-name" type="xstring" value="KNIME Views"/>
    <entry key="node-bundle-symbolic-name" type="xstring" value="org.knime.base.views"/>
    <entry key="node-bundle-vendor" type="xstring" value="KNIME AG, Zurich, Switzerland"/>
    <entry key="node-bundle-version" type="xstring" value="5.12.0"/>
    <entry key="node-feature-name" type="xstring" value="KNIME Views"/>
    <entry key="node-feature-symbolic-name" type="xstring" value="org.knime.features.base.views.feature.group"/>
    <entry key="node-feature-vendor" type="xstring" value="KNIME AG, Zurich, Switzerland"/>
    <entry key="node-feature-version" type="xstring" value="5.12.0"/>
    <config key="factory_settings"/>
    <config key="node_creation_config">
        <config key="Input Table"/>
        <config key="Output Image"/>
    </config>
    <entry key="name" type="xstring" value="Scatter Plot"/>
    <entry key="hasContent" type="xboolean" value="false"/>
    <entry key="isInactive" type="xboolean" value="false"/>
    <config key="ports"/>
    <config key="filestores">
        <entry key="file_store_location" type="xstring" isnull="true" value=""/>
        <entry key="file_store_id" type="xstring" isnull="true" value=""/>
    </config>
</config>
'''


def workflow_knime():
    def node(i, folder, x, y, w=110, h=90):
        return f'''        <config key="node_{i}">
            <entry key="id" type="xint" value="{i}"/>
            <entry key="node_settings_file" type="xstring" value="{folder}/settings.xml"/>
            <entry key="node_is_meta" type="xboolean" value="false"/>
            <entry key="node_type" type="xstring" value="NativeNode"/>
            <entry key="ui_classname" type="xstring" value="org.knime.core.node.workflow.NodeUIInformation"/>
            <config key="ui_settings">
                <config key="extrainfo.node.bounds">
                    <entry key="array-size" type="xint" value="4"/>
                    <entry key="0" type="xint" value="{x}"/>
                    <entry key="1" type="xint" value="{y}"/>
                    <entry key="2" type="xint" value="{w}"/>
                    <entry key="3" type="xint" value="{h}"/>
                </config>
            </config>
        </config>'''

    def conn(i, src, dst, sp=1, dp=1):
        return f'''        <config key="connection_{i}">
            <entry key="sourceID" type="xint" value="{src}"/>
            <entry key="destID" type="xint" value="{dst}"/>
            <entry key="sourcePort" type="xint" value="{sp}"/>
            <entry key="destPort" type="xint" value="{dp}"/>
        </config>'''

    return f'''<?xml version="1.0" encoding="UTF-8"?>
<config {NS} key="workflow.knime">
    <entry key="created_by" type="xstring" value="5.12.0"/>
    <entry key="created_by_nightly" type="xboolean" value="false"/>
    <entry key="version" type="xstring" value="5.1.0"/>
    <entry key="name" type="xstring" value="LR1_Courier_LinReg"/>
    <config key="authorInformation">
        <entry key="authored-by" type="xstring" value="Burenkov G.V."/>
        <entry key="authored-when" type="xstring" value="2026-09-25 23:30:00 +0300"/>
        <entry key="lastEdited-by" type="xstring" value="Burenkov G.V."/>
        <entry key="lastEdited-when" type="xstring" value="2026-09-25 23:30:00 +0300"/>
    </config>
    <entry key="customDescription" type="xstring" value="Lab 1 linear regression: courier distance vs delivery time"/>
    <entry key="state" type="xstring" value="IDLE"/>
    <config key="workflow_credentials"/>
    <config key="annotations">
        <config key="annotation_0">
            <entry key="text" type="xstring" value="&lt;p&gt;&lt;strong&gt;ЛР1 МИИ — линейная регрессия&lt;/strong&gt;&lt;/p&gt;&lt;p&gt;Прогноз времени доставки курьера (time_min) по расстоянию (distance_km). Обучающая выборка 20%, линейная выборка.&lt;/p&gt;"/>
            <entry key="contentType" type="xstring" value="text/html"/>
            <entry key="bgcolor" type="xint" value="16777215"/>
            <entry key="x-coordinate" type="xint" value="80"/>
            <entry key="y-coordinate" type="xint" value="40"/>
            <entry key="width" type="xint" value="980"/>
            <entry key="height" type="xint" value="90"/>
            <entry key="alignment" type="xstring" value="LEFT"/>
            <entry key="borderSize" type="xint" value="8"/>
            <entry key="borderColor" type="xint" value="16766976"/>
            <entry key="defFontSize" type="xint" value="-1"/>
            <entry key="annotation-version" type="xint" value="20230412"/>
            <config key="styles"/>
        </config>
    </config>
    <config key="nodes">
{node(1, "CSV Reader (#1)", 80, 220)}
{node(2, "Color Manager (#2)", 240, 220)}
{node(3, "Partitioning (#3)", 420, 220)}
{node(4, "Linear Regression Learner (#4)", 620, 160, 150, 100)}
{node(5, "Regression Predictor (#5)", 840, 250, 140, 100)}
{node(6, "Numeric Scorer (#6)", 1060, 220, 120, 90)}
{node(7, "Scatter Plot (#7)", 840, 420)}
    </config>
    <config key="connections">
{conn(0, 1, 2)}
{conn(1, 2, 3)}
{conn(2, 3, 4, 1, 1)}
{conn(3, 4, 5, 1, 1)}
{conn(4, 3, 5, 2, 2)}
{conn(5, 5, 6)}
{conn(6, 5, 7)}
    </config>
    <config key="workflow_editor_settings">
        <entry key="workflow.editor.snapToGrid" type="xboolean" value="true"/>
        <entry key="workflow.editor.ShowGrid" type="xboolean" value="true"/>
        <entry key="workflow.editor.gridX" type="xint" value="20"/>
        <entry key="workflow.editor.gridY" type="xint" value="20"/>
        <entry key="workflow.editor.zoomLevel" type="xdouble" value="1.0"/>
        <entry key="workflow.editor.curvedConnections" type="xboolean" value="true"/>
        <entry key="workflow.editor.connectionWidth" type="xint" value="2"/>
    </config>
</config>
'''


def workflowset_meta():
    return '''<?xml version="1.0" encoding="UTF-8"?>
<workflowset xmlns="http://www.knime.org/2008/09/XMLConfig">
</workflowset>
'''


def metadata():
    return '''<?xml version="1.0" encoding="UTF-8"?>
<workflow-metadata>
    <lastEdited>2026-09-25 23:30:00 +0300</lastEdited>
    <description>Linear regression: courier delivery time vs distance</description>
</workflow-metadata>
'''


def write_wf(root: Path):
    if root.exists():
        shutil.rmtree(root)
    nodes = {
        "CSV Reader (#1)": csv_reader(),
        "Color Manager (#2)": color_manager(),
        "Partitioning (#3)": partitioning(),
        "Linear Regression Learner (#4)": linreg(),
        "Regression Predictor (#5)": predictor(),
        "Numeric Scorer (#6)": scorer(),
        "Scatter Plot (#7)": scatter(),
    }
    for folder, xml in nodes.items():
        d = root / folder
        d.mkdir(parents=True, exist_ok=True)
        (d / "settings.xml").write_text(xml, encoding="utf-8")
    (root / "workflow.knime").write_text(workflow_knime(), encoding="utf-8")
    (root / "workflowset.meta").write_text(workflowset_meta(), encoding="utf-8")
    (root / "workflow-metadata.xml").write_text(metadata(), encoding="utf-8")
    data = root / "data"
    data.mkdir(exist_ok=True)
    shutil.copy2(CSV_SRC, data / "courier_delivery.csv")
    print("wrote", root)


def main():
    write_wf(WF)
    try:
        write_wf(WS_WF)
    except Exception as e:
        print("workspace copy failed", e)


if __name__ == "__main__":
    main()
