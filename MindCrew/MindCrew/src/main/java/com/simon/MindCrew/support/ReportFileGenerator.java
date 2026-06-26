package com.simon.MindCrew.support;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 报告文件生成器（提示词驱动）。
 *
 * - {@link #isReportRequest(String)} 判断用户是否在要"报告"。
 * - {@link #detectFormat(String)}    判断要 docx 还是 txt。
 * - {@link #toTxt(String)}           Markdown → 纯文本。
 * - {@link #toDocx(String, String)}  Markdown → Word（POI），支持标题/段落/加粗/列表/表格。
 *
 * mermaid / echarts 代码块在 docx/txt 里降级为占位说明（图表不嵌入文件）。
 */
@Slf4j
@Component
public class ReportFileGenerator {

    /** 产出类动词 + 报告/word/文档/文件 等，即认为是"生成可下载文件"的诉求（放宽匹配，不再死盯"报告"二字） */
    private static final Pattern REPORT_PATTERN = Pattern.compile(
            "(生成|导出|输出|下载|做成|做一?份|做个|整理成|整理为|写一?份|写个|出一?份|给我).{0,8}(报告|report|word|文档|文件|doc|docx|txt)"
            + "|(报告|report|word|文档|文件).{0,6}(下载|导出|生成|文件)",
            Pattern.CASE_INSENSITIVE);

    private static final Pattern DOCX_HINT = Pattern.compile(
            "word|文档|docx|\\.doc|办公", Pattern.CASE_INSENSITIVE);

    /** 行内加粗 **xxx** */
    private static final Pattern BOLD = Pattern.compile("\\*\\*(.+?)\\*\\*");

    public boolean isReportRequest(String question) {
        if (question == null || question.isBlank()) return false;
        return REPORT_PATTERN.matcher(question).find();
    }

    /** 返回 "docx" 或 "txt"，默认 txt */
    public String detectFormat(String question) {
        if (question != null && DOCX_HINT.matcher(question).find()) {
            return "docx";
        }
        return "txt";
    }

    // ====================================================================
    // TXT
    // ====================================================================

    /** 把 Markdown 剥成可读纯文本 */
    public String toTxt(String markdown) {
        if (markdown == null) return "";
        StringBuilder out = new StringBuilder();
        boolean inFence = false;
        String fenceLang = "";
        for (String raw : markdown.split("\n", -1)) {
            String line = raw;
            String trimmed = line.trim();

            // 代码围栏
            if (trimmed.startsWith("```")) {
                if (!inFence) {
                    inFence = true;
                    fenceLang = trimmed.substring(3).trim().toLowerCase();
                    if (isDiagramLang(fenceLang)) {
                        out.append("【图表：").append(fenceLang).append("，详见在线版本】\n");
                    }
                } else {
                    inFence = false;
                    fenceLang = "";
                }
                continue;
            }
            if (inFence) {
                if (!isDiagramLang(fenceLang)) out.append(line).append("\n");
                continue;
            }

            // 标题
            Matcher h = HEADING.matcher(line);
            if (h.matches()) {
                out.append(stripInline(h.group(2))).append("\n");
                continue;
            }
            // 表格分隔行 |:---|---| 跳过
            if (TABLE_SEP.matcher(trimmed).matches()) continue;
            // 表格行 → 用制表符分隔
            if (trimmed.startsWith("|") && trimmed.endsWith("|")) {
                String[] cells = splitTableRow(trimmed);
                out.append(String.join("\t", stripCells(cells))).append("\n");
                continue;
            }
            // 列表
            Matcher ul = UL.matcher(line);
            if (ul.matches()) {
                out.append("· ").append(stripInline(ul.group(2))).append("\n");
                continue;
            }
            Matcher ol = OL.matcher(line);
            if (ol.matches()) {
                out.append(ol.group(1)).append(". ").append(stripInline(ol.group(2))).append("\n");
                continue;
            }
            out.append(stripInline(line)).append("\n");
        }
        return out.toString();
    }

    // ====================================================================
    // DOCX
    // ====================================================================

    private static final Pattern HEADING   = Pattern.compile("^(#{1,6})\\s+(.*)$");
    private static final Pattern UL         = Pattern.compile("^\\s*[-*+]\\s+(\\s*)?(.*)$");
    private static final Pattern OL         = Pattern.compile("^\\s*(\\d+)[.)]\\s+(.*)$");
    private static final Pattern TABLE_SEP  = Pattern.compile("^\\|?\\s*:?-{2,}.*$");

    /** Markdown → docx 字节数组 */
    public byte[] toDocx(String markdown, String title) {
        try (XWPFDocument doc = new XWPFDocument();
             ByteArrayOutputStream bos = new ByteArrayOutputStream()) {

            if (title != null && !title.isBlank()) {
                XWPFParagraph tp = doc.createParagraph();
                tp.setAlignment(ParagraphAlignment.CENTER);
                XWPFRun tr = tp.createRun();
                tr.setText(title);
                tr.setBold(true);
                tr.setFontSize(20);
            }

            String[] lines = (markdown == null ? "" : markdown).split("\n", -1);
            int i = 0;
            boolean inFence = false;
            String fenceLang = "";
            while (i < lines.length) {
                String line = lines[i];
                String trimmed = line.trim();

                // 代码围栏
                if (trimmed.startsWith("```")) {
                    if (!inFence) {
                        inFence = true;
                        fenceLang = trimmed.substring(3).trim().toLowerCase();
                        if (isDiagramLang(fenceLang)) {
                            XWPFParagraph p = doc.createParagraph();
                            XWPFRun r = p.createRun();
                            r.setItalic(true);
                            r.setColor("888888");
                            r.setText("【图表：" + fenceLang + "，请在在线版本查看渲染效果】");
                        }
                    } else {
                        inFence = false;
                        fenceLang = "";
                    }
                    i++;
                    continue;
                }
                if (inFence) {
                    if (!isDiagramLang(fenceLang)) {
                        XWPFParagraph p = doc.createParagraph();
                        XWPFRun r = p.createRun();
                        r.setFontFamily("Consolas");
                        r.setText(line);
                    }
                    i++;
                    continue;
                }

                // 表格：连续多行 | ... |
                if (trimmed.startsWith("|") && trimmed.endsWith("|")) {
                    List<String[]> tableRows = new ArrayList<>();
                    while (i < lines.length) {
                        String t = lines[i].trim();
                        if (!(t.startsWith("|") && t.endsWith("|"))) break;
                        if (!isSeparatorRow(t)) {
                            tableRows.add(stripCells(splitTableRow(t)));
                        }
                        i++;
                    }
                    if (!tableRows.isEmpty()) writeTable(doc, tableRows);
                    continue;
                }

                // 标题
                Matcher h = HEADING.matcher(line);
                if (h.matches()) {
                    int level = h.group(1).length();
                    XWPFParagraph p = doc.createParagraph();
                    XWPFRun r = p.createRun();
                    r.setBold(true);
                    r.setFontSize(Math.max(11, 18 - (level - 1) * 2));
                    r.setText(stripInline(h.group(2)));
                    i++;
                    continue;
                }

                // 无序列表
                Matcher ul = UL.matcher(line);
                if (ul.matches()) {
                    XWPFParagraph p = doc.createParagraph();
                    p.setIndentationLeft(360);
                    XWPFRun bullet = p.createRun();
                    bullet.setText("• ");
                    writeInline(p, ul.group(2));
                    i++;
                    continue;
                }
                // 有序列表
                Matcher ol = OL.matcher(line);
                if (ol.matches()) {
                    XWPFParagraph p = doc.createParagraph();
                    p.setIndentationLeft(360);
                    XWPFRun num = p.createRun();
                    num.setText(ol.group(1) + ". ");
                    writeInline(p, ol.group(2));
                    i++;
                    continue;
                }

                // 普通段落（空行也产出一个空段，保留间距）
                XWPFParagraph p = doc.createParagraph();
                if (!trimmed.isEmpty()) writeInline(p, line);
                i++;
            }

            doc.write(bos);
            return bos.toByteArray();
        } catch (IOException e) {
            log.error("[ReportFileGenerator] docx 生成失败", e);
            throw new RuntimeException("报告生成失败: " + e.getMessage());
        }
    }

    // ====================================================================
    // helpers
    // ====================================================================

    private boolean isDiagramLang(String lang) {
        return "mermaid".equals(lang) || "echarts".equals(lang);
    }

    private boolean isSeparatorRow(String trimmedRow) {
        for (String c : splitTableRow(trimmedRow)) {
            String v = c.trim();
            if (!v.isEmpty() && !v.matches(":?-{1,}:?")) return false;
        }
        return true;
    }

    private String[] splitTableRow(String row) {
        String body = row.trim();
        if (body.startsWith("|")) body = body.substring(1);
        if (body.endsWith("|")) body = body.substring(0, body.length() - 1);
        return body.split("\\|", -1);
    }

    private String[] stripCells(String[] cells) {
        String[] out = new String[cells.length];
        for (int i = 0; i < cells.length; i++) out[i] = stripInline(cells[i].trim());
        return out;
    }

    /** 写一个带表头加粗的 Word 表格 */
    private void writeTable(XWPFDocument doc, List<String[]> rows) {
        int cols = 0;
        for (String[] r : rows) cols = Math.max(cols, r.length);
        if (cols == 0) return;

        XWPFTable table = doc.createTable(rows.size(), cols);
        table.setWidth("100%");
        for (int r = 0; r < rows.size(); r++) {
            XWPFTableRow row = table.getRow(r);
            String[] cells = rows.get(r);
            for (int c = 0; c < cols; c++) {
                XWPFTableCell cell = row.getCell(c);
                XWPFParagraph p = cell.getParagraphs().isEmpty()
                        ? cell.addParagraph() : cell.getParagraphs().get(0);
                XWPFRun run = p.createRun();
                run.setText(c < cells.length ? cells[c] : "");
                if (r == 0) run.setBold(true);
            }
        }
        doc.createParagraph(); // 表格后空一行
    }

    /** 把含 **加粗** 的行写进段落，加粗片段单独成 run */
    private void writeInline(XWPFParagraph p, String text) {
        Matcher m = BOLD.matcher(text);
        int last = 0;
        while (m.find()) {
            if (m.start() > last) {
                XWPFRun r = p.createRun();
                r.setText(text.substring(last, m.start()));
            }
            XWPFRun b = p.createRun();
            b.setBold(true);
            b.setText(m.group(1));
            last = m.end();
        }
        if (last < text.length()) {
            XWPFRun r = p.createRun();
            r.setText(text.substring(last));
        }
    }

    /** 去掉行内 Markdown 记号（加粗/行内代码/链接），返回可读文本 */
    private String stripInline(String text) {
        if (text == null) return "";
        String s = BOLD.matcher(text).replaceAll("$1");
        s = s.replaceAll("`([^`]*)`", "$1");
        s = s.replaceAll("\\[([^\\]]*)\\]\\([^)]*\\)", "$1"); // [text](url) -> text
        return s.trim();
    }
}
