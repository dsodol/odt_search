import javax.swing.*;
import javax.swing.event.TreeExpansionEvent;
import javax.swing.event.TreeWillExpandListener;
import javax.swing.event.TreeSelectionListener;
import javax.swing.table.DefaultTableModel;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreePath;
import javax.swing.tree.TreeSelectionModel;
import javax.swing.tree.ExpandVetoException;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.List;
import java.util.*;
import java.util.concurrent.CancellationException;
import java.util.prefs.Preferences;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.awt.image.BufferedImage;

import org.w3c.dom.*;
import org.xml.sax.InputSource;

import javax.imageio.ImageIO;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

/**
 * ODT Search UI Tool
 * - Directory tree to select a directory
 * - Search term input field
 * - Search result area
 * - Separate window to view the search log
 *
 * No external libraries are required. ODT files are read as ZIP archives and their content.xml is parsed via JAXP.
 */
public class OdtSearchApp extends JFrame {
    private final JComboBox<String> searchCombo = new JComboBox<>();
    private final JButton searchButton = new JButton("Search");
    private final JButton stopButton = new JButton("Stop");
    private final JButton chooseRootButton = new JButton("Choose Root…");
    private final JButton openLogButton = new JButton("Open Log Window");
    private final JLabel statusLabel = new JLabel("Ready");

    private final JTable resultsTable = new JTable();
    private final DefaultTableModel resultsModel = new DefaultTableModel(new Object[]{"File", "Matches", "Snippet"}, 0) {
        @Override
        public boolean isCellEditable(int row, int column) { return false; }
    };
    // Track files already shown in the results to avoid duplicates
    private final Map<String, Integer> fileRowIndex = new HashMap<>();

    // Search option controls
    private final JCheckBox exactPhraseCheck = new JCheckBox("Exact phrase");
    private final JCheckBox ignoreSpacesCheck = new JCheckBox("Ignore spaces");
    private final JCheckBox proximityCheck = new JCheckBox("Proximity");
    private final JSpinner proximitySpinner = new JSpinner(new SpinnerNumberModel(3, 1, 10, 1));

    private final JTree fileTree = new JTree();
    private Path rootPath = Paths.get(System.getProperty("user.home"));

    // Preferences and search history
    private final Preferences prefs = Preferences.userNodeForPackage(OdtSearchApp.class);
    private static final String PREF_LAST_ROOT = "lastRoot";
    private static final String PREF_LAST_SELECTED = "lastSelectedDir";
    private static final String PREF_HISTORY = "searchHistory";
    private static final String PREF_OPT_EXACT = "optExactPhrase";
    private static final String PREF_OPT_IGNORE_SPACES = "optIgnoreSpaces";
    private static final String PREF_OPT_PROX = "optProximity";
    private static final String PREF_OPT_PROX_N = "optProximityN";
    private final java.util.List<String> searchHistory = new ArrayList<>();

    private SearchWorker currentWorker;
    private final LogWindow logWindow = new LogWindow();

    // Search options data holder
    private static class SearchOptions {
        final boolean exactPhrase;
        final boolean ignoreSpaces;
        final boolean proximity;
        final int proximityN;
        SearchOptions(boolean exactPhrase, boolean ignoreSpaces, boolean proximity, int proximityN) {
            this.exactPhrase = exactPhrase;
            this.ignoreSpaces = ignoreSpaces;
            this.proximity = proximity;
            this.proximityN = proximityN;
        }
        String summary() {
            if (proximity) return "proximity<=" + proximityN;
            if (exactPhrase && ignoreSpaces) return "exact phrase (ignore spaces)";
            if (exactPhrase) return "exact phrase";
            return "substring";
        }
    }

    public OdtSearchApp() {
        super("Document Search Tool");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1100, 700);
        setLocationRelativeTo(null);

        initUI();
        initActions();

        // Load preferences: last root and history
        try {
            String lastRoot = prefs.get(PREF_LAST_ROOT, null);
            if (lastRoot != null) {
                Path p = Paths.get(lastRoot);
                if (Files.isDirectory(p)) {
                    rootPath = p;
                } else {
                    // Clear invalid stored path to avoid confusion next time
                    try { prefs.remove(PREF_LAST_ROOT); prefs.flush(); } catch (Exception ignore2) {}
                }
            }
            loadHistoryFromPrefs();
            refreshSearchComboModel();
            // Load search options
            exactPhraseCheck.setSelected(prefs.getBoolean(PREF_OPT_EXACT, true));
            ignoreSpacesCheck.setSelected(prefs.getBoolean(PREF_OPT_IGNORE_SPACES, false));
            boolean prox = prefs.getBoolean(PREF_OPT_PROX, false);
            proximityCheck.setSelected(prox);
            int pn = Math.max(1, Math.min(10, prefs.getInt(PREF_OPT_PROX_N, 3)));
            proximitySpinner.setValue(pn);
            proximitySpinner.setEnabled(prox);
        } catch (Exception ignore) {
        }

        // Set application icon images (multi-size) generated programmatically
        try {
            setIconImages(IconFactory.createAppIcons());
        } catch (Exception ex) {
            // Fallback silently; app will still run without custom icon
        }

        Path toSet = rootPath; // capture effectively final
        SwingUtilities.invokeLater(() -> setRoot(toSet));
    }

    private void initUI() {
        // Top panel: controls
        JPanel controls = new JPanel(new BorderLayout(8, 8));
        JPanel leftControls = new JPanel(new FlowLayout(FlowLayout.LEFT));
        leftControls.add(new JLabel("Search term:"));
        searchCombo.setEditable(true);
        // Set a reasonable width for the editor component
        Component editorComp = searchCombo.getEditor().getEditorComponent();
        if (editorComp instanceof JTextField) {
            ((JTextField) editorComp).setColumns(30);
        }
        leftControls.add(searchCombo);
        leftControls.add(searchButton);
        leftControls.add(stopButton);
        stopButton.setEnabled(false);

        // Options panel
        JPanel optionsPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        exactPhraseCheck.setToolTipText("Match the phrase exactly (case-insensitive)");
        ignoreSpacesCheck.setToolTipText("Treat whitespace as optional");
        proximityCheck.setToolTipText("Match words within N-word distance, in order");
        proximitySpinner.setToolTipText("Max distance between adjacent words");
        optionsPanel.add(exactPhraseCheck);
        optionsPanel.add(ignoreSpacesCheck);
        optionsPanel.add(proximityCheck);
        optionsPanel.add(new JLabel("N:"));
        optionsPanel.add(proximitySpinner);
        proximitySpinner.setEnabled(false);

        JPanel leftStack = new JPanel();
        leftStack.setLayout(new BoxLayout(leftStack, BoxLayout.Y_AXIS));
        leftStack.add(leftControls);
        leftStack.add(optionsPanel);

        JPanel rightControls = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        rightControls.add(chooseRootButton);
        rightControls.add(openLogButton);

        controls.add(leftStack, BorderLayout.WEST);
        controls.add(rightControls, BorderLayout.EAST);

        // Left panel: directory tree
        fileTree.setModel(new DefaultTreeModel(new DefaultMutableTreeNode("loading…")));
        fileTree.getSelectionModel().setSelectionMode(TreeSelectionModel.SINGLE_TREE_SELECTION);
        fileTree.setRootVisible(true);
        fileTree.addTreeWillExpandListener(new LazyLoader());
        JScrollPane treeScroll = new JScrollPane(fileTree);

        // Right panel: results table
        resultsTable.setModel(resultsModel);
        resultsTable.setAutoCreateRowSorter(true);
        resultsTable.setFillsViewportHeight(true);
        resultsTable.setRowHeight(22);
        resultsTable.getColumnModel().getColumn(0).setPreferredWidth(520);
        resultsTable.getColumnModel().getColumn(1).setPreferredWidth(80);
        resultsTable.getColumnModel().getColumn(2).setPreferredWidth(400);
        JScrollPane resultsScroll = new JScrollPane(resultsTable);

        JSplitPane split = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, treeScroll, resultsScroll);
        split.setDividerLocation(350);

        // Status bar
        JPanel status = new JPanel(new BorderLayout());
        status.setBorder(BorderFactory.createEmptyBorder(4, 8, 4, 8));
        status.add(statusLabel, BorderLayout.WEST);

        // Layout frame
        getContentPane().setLayout(new BorderLayout(8, 8));
        getContentPane().add(controls, BorderLayout.NORTH);
        getContentPane().add(split, BorderLayout.CENTER);
        getContentPane().add(status, BorderLayout.SOUTH);
    }

    private void initActions() {
        chooseRootButton.addActionListener(this::onChooseRoot);
        searchButton.addActionListener(this::onStartSearch);
        stopButton.addActionListener(this::onStopSearch);
        openLogButton.addActionListener(e -> logWindow.setVisible(true));

        // Start search when user presses Enter in the combo editor OR selects an item.
        // Note: Only attach to the combo box itself. Attaching a listener to the editor's
        // JTextField as well causes duplicate ActionEvents and triggers the "A search is already running" warning.
        searchCombo.addActionListener(this::onStartSearch);

        // Enable/disable proximity spinner based on checkbox
        proximityCheck.addItemListener(e2 -> proximitySpinner.setEnabled(proximityCheck.isSelected()));

        // Persist last selected directory when tree selection changes
        fileTree.addTreeSelectionListener(e -> {
            TreePath sel = fileTree.getSelectionPath();
            if (sel == null) return;
            Object last = ((DefaultMutableTreeNode) sel.getLastPathComponent()).getUserObject();
            if (last instanceof FileNode) {
                Path p = ((FileNode) last).path;
                try {
                    if (Files.isDirectory(p)) {
                        prefs.put(PREF_LAST_SELECTED, p.toAbsolutePath().toString());
                        prefs.flush();
                    }
                } catch (Exception ignore) {}
            }
        });

        // Double-click a result row to open the document
        resultsTable.addMouseListener(new MouseAdapter() {
            @Override public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2 && SwingUtilities.isLeftMouseButton(e)) {
                    int row = resultsTable.rowAtPoint(e.getPoint());
                    if (row >= 0) {
                        int modelRow = resultsTable.convertRowIndexToModel(row);
                        String path = (String) resultsModel.getValueAt(modelRow, 0);
                        try {
                            Desktop.getDesktop().open(new File(path));
                        } catch (Exception ex) {
                            JOptionPane.showMessageDialog(OdtSearchApp.this, "Cannot open file:\n" + path + "\n" + ex.getMessage(), "Open Error", JOptionPane.ERROR_MESSAGE);
                        }
                    }
                }
            }
        });

        addWindowListener(new WindowAdapter() {
            @Override public void windowClosing(WindowEvent e) {
                if (currentWorker != null && !currentWorker.isDone()) currentWorker.cancel(true);
                try { prefs.flush(); } catch (Exception ex1) {
                    try { prefs.sync(); } catch (Exception ex2) { /* ignore */ }
                }
            }
        });
    }

    private void onChooseRoot(ActionEvent e) {
        JFileChooser chooser = new JFileChooser(rootPath.toFile());
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.setDialogTitle("Select Root Directory");
        int res = chooser.showOpenDialog(this);
        if (res == JFileChooser.APPROVE_OPTION) {
            setRoot(chooser.getSelectedFile().toPath());
        }
    }

    private void setRoot(Path newRoot) {
        this.rootPath = newRoot;
        // persist last directory
        try {
            prefs.put(PREF_LAST_ROOT, rootPath.toAbsolutePath().toString());
            prefs.flush();
        } catch (Exception ignore) {}
        setTitle("Document Search Tool — " + rootPath.toAbsolutePath());
        logWindow.log("Root set to: " + rootPath);
        DefaultMutableTreeNode rootNode = new DefaultMutableTreeNode(new FileNode(rootPath));
        // Eagerly populate root children so the tree isn't stuck showing just a placeholder
        loadChildren(rootNode);
        fileTree.setModel(new DefaultTreeModel(rootNode));
        ((DefaultTreeModel) fileTree.getModel()).reload();
        TreePath rootTreePath = new TreePath(rootNode.getPath());
        fileTree.expandPath(rootTreePath);

        // Try to restore last selected directory (under this root)
        restoreLastSelectedDirectory();
    }

    // ===== Search history helpers =====
    private void loadHistoryFromPrefs() {
        searchHistory.clear();
        String raw = prefs.get(PREF_HISTORY, "");
        if (raw != null && !raw.isEmpty()) {
            String[] items = raw.split("\n");
            for (String s : items) {
                String term = s.trim();
                if (!term.isEmpty()) searchHistory.add(term);
            }
        }
    }

    private void saveHistoryToPrefs() {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < searchHistory.size(); i++) {
            if (i > 0) sb.append('\n');
            // sanitize newlines just in case
            sb.append(searchHistory.get(i).replace('\n', ' '));
        }
        try { prefs.put(PREF_HISTORY, sb.toString()); } catch (Exception ignore) {}
    }

    private void refreshSearchComboModel() {
        DefaultComboBoxModel<String> model = new DefaultComboBoxModel<>(searchHistory.toArray(new String[0]));
        searchCombo.setModel(model);
        searchCombo.setEditable(true);
    }

    private void addToHistory(String term) {
        // move to front, unique, cap at 20
        searchHistory.remove(term);
        searchHistory.add(0, term);
        while (searchHistory.size() > 20) searchHistory.remove(searchHistory.size() - 1);
    }


    private void onStartSearch(ActionEvent e) {
        if (currentWorker != null && !currentWorker.isDone()) {
            JOptionPane.showMessageDialog(this, "A search is already running.", "Busy", JOptionPane.WARNING_MESSAGE);
            return;
        }
        String term = "";
        Component ed = searchCombo.getEditor().getEditorComponent();
        if (ed instanceof JTextField) term = ((JTextField) ed).getText().trim();
        else if (searchCombo.getSelectedItem() != null) term = searchCombo.getSelectedItem().toString().trim();
        if (term.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Enter a search term.", "Validation", JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        // maintain search history and refresh combo model
        addToHistory(term);
        saveHistoryToPrefs();
        refreshSearchComboModel();
        searchCombo.setSelectedItem(term);

        // Determine selected directory from tree; if none, use root
        Path startDir = rootPath;
        TreePath sel = fileTree.getSelectionPath();
        if (sel != null) {
            Object last = ((DefaultMutableTreeNode) sel.getLastPathComponent()).getUserObject();
            if (last instanceof FileNode) {
                startDir = ((FileNode) last).path;
            }
        }

        resultsModel.setRowCount(0);
        fileRowIndex.clear();
        searchButton.setEnabled(false);
        stopButton.setEnabled(true);
        // Build a short status with options summary
        statusLabel.setText("Searching…");
        logWindow.log("Starting search for '" + term + "' in " + startDir);

        // Build search options
        boolean optProx = proximityCheck.isSelected();
        boolean optExact = exactPhraseCheck.isSelected();
        boolean optIgnoreSpaces = ignoreSpacesCheck.isSelected();
        int proxN = (int) proximitySpinner.getValue();
        if (optProx) { // proximity overrides other modes
            optExact = false;
            optIgnoreSpaces = false;
        }
        SearchOptions options = new SearchOptions(optExact, optIgnoreSpaces, optProx, proxN);
        // persist options
        try {
            prefs.putBoolean(PREF_OPT_EXACT, optExact);
            prefs.putBoolean(PREF_OPT_IGNORE_SPACES, optIgnoreSpaces);
            prefs.putBoolean(PREF_OPT_PROX, optProx);
            prefs.putInt(PREF_OPT_PROX_N, proxN);
            prefs.flush();
        } catch (Exception ignore) {}

        // Disable option controls during search
        exactPhraseCheck.setEnabled(false);
        ignoreSpacesCheck.setEnabled(false);
        proximityCheck.setEnabled(false);
        proximitySpinner.setEnabled(false);

        logWindow.log("Options: " + options.summary());

        currentWorker = new SearchWorker(startDir, term, options, this::onResult, this::onProgress, this::onCompleted, logWindow);
        currentWorker.execute();
    }

    private void onStopSearch(ActionEvent e) {
        if (currentWorker != null && !currentWorker.isDone()) {
            currentWorker.cancel(true);
            stopButton.setEnabled(false);
            statusLabel.setText("Cancelling…");
            logWindow.log("Cancellation requested by user.");
        }
    }

    private void onResult(SearchResult r) {
        String pathStr = r.file.toAbsolutePath().toString();
        String osName = System.getProperty("os.name");
        if (osName != null && osName.toLowerCase(Locale.ROOT).contains("win")) {
            pathStr = pathStr.toLowerCase(Locale.ROOT);
        }
        Integer existingRow = fileRowIndex.get(pathStr);
        if (existingRow != null) {
            // Update the existing row: accumulate match count, keep the first snippet
            Object oldVal = resultsModel.getValueAt(existingRow, 1);
            int oldCount = 0;
            if (oldVal instanceof Number) oldCount = ((Number) oldVal).intValue();
            resultsModel.setValueAt(oldCount + r.matchCount, existingRow, 1);
        } else {
            resultsModel.addRow(new Object[]{r.file.toString(), r.matchCount, r.snippet});
            int newRow = resultsModel.getRowCount() - 1;
            fileRowIndex.put(pathStr, newRow);
        }
    }

    private void onProgress(String message) {
        statusLabel.setText(message);
    }

    private void onCompleted(Throwable error, boolean cancelled, long filesScanned, long matchesFound, long elapsedMillis) {
        searchButton.setEnabled(true);
        stopButton.setEnabled(false);
        // Re-enable option controls
        exactPhraseCheck.setEnabled(true);
        ignoreSpacesCheck.setEnabled(true);
        proximityCheck.setEnabled(true);
        proximitySpinner.setEnabled(proximityCheck.isSelected());
        if (cancelled) {
            statusLabel.setText("Cancelled. Scanned " + filesScanned + " files, found " + matchesFound + " matches.");
            logWindow.log("Search cancelled. Files scanned: " + filesScanned + ", matches: " + matchesFound + ".");
        } else if (error != null) {
            statusLabel.setText("Error: " + error.getMessage());
            logWindow.log("Error: " + getStackTrace(error));
        } else {
            statusLabel.setText("Done. Scanned " + filesScanned + " files, found " + matchesFound + " matches. In " + (elapsedMillis/1000.0) + "s");
            logWindow.log("Search completed. Files scanned: " + filesScanned + ", matches: " + matchesFound + ".");
        }
    }

    private static String getStackTrace(Throwable t) {
        StringWriter sw = new StringWriter();
        t.printStackTrace(new PrintWriter(sw));
        return sw.toString();
    }

    // Populate a directory node with its immediate subdirectories (adds a placeholder child for lazy loading)
    private void loadChildren(DefaultMutableTreeNode node) {
        Object uo = node.getUserObject();
        if (!(uo instanceof FileNode)) return;
        FileNode fn = (FileNode) uo;
        try (DirectoryStream<Path> ds = Files.newDirectoryStream(fn.path)) {
            java.util.List<DefaultMutableTreeNode> children = new ArrayList<>();
            for (Path p : ds) {
                if (Files.isDirectory(p)) {
                    DefaultMutableTreeNode child = new DefaultMutableTreeNode(new FileNode(p));
                    // add a dummy child so it can be expanded lazily later
                    child.add(new DefaultMutableTreeNode("loading"));
                    children.add(child);
                }
            }
            children.sort(Comparator.comparing(a -> ((FileNode) a.getUserObject()).path.getFileName().toString().toLowerCase()));
            for (DefaultMutableTreeNode c : children) node.add(c);
        } catch (IOException ex) {
            logWindow.log("Failed to list directory: " + ex.getMessage());
        }
    }

    // Lazy loading directory tree
    private class LazyLoader implements TreeWillExpandListener {
        @Override
        public void treeWillExpand(TreeExpansionEvent event) throws ExpandVetoException {
            DefaultMutableTreeNode node = (DefaultMutableTreeNode) event.getPath().getLastPathComponent();
            Object uo = node.getUserObject();
            if (!(uo instanceof FileNode)) return;
            if (node.getChildCount() == 1 && ((DefaultMutableTreeNode) node.getFirstChild()).getUserObject() instanceof String) {
                // load children
                node.removeAllChildren();
                FileNode fn = (FileNode) uo;
                try (DirectoryStream<Path> ds = Files.newDirectoryStream(fn.path)) {
                    List<DefaultMutableTreeNode> children = new ArrayList<>();
                    for (Path p : ds) {
                        if (Files.isDirectory(p)) {
                            DefaultMutableTreeNode child = new DefaultMutableTreeNode(new FileNode(p));
                            child.add(new DefaultMutableTreeNode("loading"));
                            children.add(child);
                        }
                    }
                    children.sort(Comparator.comparing(a -> ((FileNode) a.getUserObject()).path.getFileName().toString().toLowerCase()));
                    for (DefaultMutableTreeNode c : children) node.add(c);
                } catch (IOException ex) {
                    logWindow.log("Failed to list directory: " + ex.getMessage());
                }
                ((DefaultTreeModel) fileTree.getModel()).reload(node);
            }
        }

        @Override
        public void treeWillCollapse(TreeExpansionEvent event) {}
    }

    private static class FileNode {
        final Path path;
        FileNode(Path path) { this.path = path; }
        @Override public String toString() { return path.getFileName() == null ? path.toString() : path.getFileName().toString(); }
    }

    // Search worker
    private static class SearchWorker extends SwingWorker<Void, SearchResult> {
        private final Path startDir;
        private final String termLower;
        private final SearchOptions options;
        private final java.util.function.Consumer<SearchResult> onResult;
        private final java.util.function.Consumer<String> onProgress;
        private final CompletionCallback onCompleted;
        private final LogWindow logger;

        private long filesScanned = 0;
        private long matchesFound = 0;
        private long lastProgressUpdate = 0L;
        private final long startTime = System.currentTimeMillis();

        SearchWorker(Path startDir, String term, SearchOptions options,
                     java.util.function.Consumer<SearchResult> onResult,
                     java.util.function.Consumer<String> onProgress,
                     CompletionCallback onCompleted, LogWindow logger) {
            this.startDir = startDir;
            this.termLower = term.toLowerCase(Locale.ROOT);
            this.options = options;
            this.onResult = onResult;
            this.onProgress = onProgress;
            this.onCompleted = onCompleted;
            this.logger = logger;
        }

        @Override
        protected Void doInBackground() {
            try {
                Files.walkFileTree(startDir, new SimpleFileVisitor<Path>() {
                    @Override
                    public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
                        if (isCancelled()) return FileVisitResult.TERMINATE;
                        String name = file.toString().toLowerCase(Locale.ROOT);
                        try {
                            if (name.endsWith(".odt")) {
                                filesScanned++;
                                String text = OdtTextExtractor.extractText(file);
                                handleTextResult(file, text);
                            } else if (name.endsWith(".docx")) {
                                filesScanned++;
                                String text = DocxTextExtractor.extractText(file);
                                handleTextResult(file, text);
                            } else if (name.endsWith(".xlsx")) {
                                filesScanned++;
                                String text = XlsxTextExtractor.extractText(file);
                                handleTextResult(file, text);
                            }
                        } catch (Exception ex) {
                            logger.log("Failed to read " + file + ": " + ex.getMessage());
                        }
                        maybeUpdateProgress();
                        return isCancelled() ? FileVisitResult.TERMINATE : FileVisitResult.CONTINUE;
                    }

                    @Override
                    public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) {
                        return isCancelled() ? FileVisitResult.TERMINATE : FileVisitResult.CONTINUE;
                    }

                    @Override
                    public FileVisitResult visitFileFailed(Path file, IOException exc) {
                        logger.log("Cannot access: " + file + " — " + (exc != null ? exc.getMessage() : "unknown error"));
                        return isCancelled() ? FileVisitResult.TERMINATE : FileVisitResult.CONTINUE;
                    }
                });
            } catch (CancellationException ce) {
                // ignore
            } catch (IOException ioe) {
                logger.log("Traversal error: " + ioe.getMessage());
            }
            return null;
        }

        private void maybeUpdateProgress() {
            long now = System.currentTimeMillis();
            if (now - lastProgressUpdate > 400) {
                lastProgressUpdate = now;
                onProgress.accept("Scanned " + filesScanned + " file(s), found " + matchesFound + " matches…");
            }
        }

        private void handleTextResult(Path file, String text) {
            if (text == null) return;
            MatchResult mr = computeMatches(text, termLower, options);
            if (mr.count > 0) {
                matchesFound += mr.count;
                String snippet = makeSnippet(text, mr.firstStart, mr.firstEnd, 200);
                publish(new SearchResult(file, mr.count, snippet));
                logger.log("Match in: " + file + " (" + mr.count + ")");
            }
        }

        @Override
        protected void process(List<SearchResult> chunks) {
            for (SearchResult r : chunks) onResult.accept(r);
        }

        @Override
        protected void done() {
            boolean cancelled = isCancelled();
            Throwable error = null;
            try {
                get();
            } catch (CancellationException ce) {
                cancelled = true;
            } catch (Exception ex) {
                error = ex.getCause() != null ? ex.getCause() : ex;
            }
            long elapsed = System.currentTimeMillis() - startTime;
            onCompleted.onCompleted(error, cancelled, filesScanned, matchesFound, elapsed);
        }

        private static int countOccurrences(String haystack, String needleLower) {
            if (haystack == null || haystack.isEmpty()) return 0;
            String h = haystack.toLowerCase(Locale.ROOT);
            int count = 0, idx = 0;
            while ((idx = h.indexOf(needleLower, idx)) >= 0) {
                count++;
                idx += needleLower.length();
            }
            return count;
        }

        // Result of matching: total count and the first match range [firstStart, firstEnd)
        private static class MatchResult {
            final int count;
            final int firstStart;
            final int firstEnd;
            MatchResult(int count, int firstStart, int firstEnd) {
                this.count = count;
                this.firstStart = firstStart;
                this.firstEnd = firstEnd;
            }
        }

        private static class Token {
            final String word;
            final int start;
            final int end;
            Token(String word, int start, int end) {
                this.word = word;
                this.start = start;
                this.end = end;
            }
        }

        private static java.util.List<Token> tokenize(String textLower) {
            java.util.List<Token> out = new ArrayList<>();
            int n = textLower.length();
            int i = 0;
            while (i < n) {
                // skip non-word
                while (i < n && !Character.isLetterOrDigit(textLower.charAt(i))) i++;
                if (i >= n) break;
                int start = i;
                while (i < n && Character.isLetterOrDigit(textLower.charAt(i))) i++;
                int end = i;
                String w = textLower.substring(start, end);
                out.add(new Token(w, start, end));
            }
            return out;
        }

        private static MatchResult computeMatches(String text, String termLower, SearchOptions options) {
            if (text == null || termLower == null || termLower.isEmpty()) return new MatchResult(0, -1, -1);
            String lower = text.toLowerCase(Locale.ROOT);
            if (options.proximity) {
                // Proximity search: words within N in order
                String[] qwordsArr = termLower.split("[^\\p{L}\\p{N}]+");
                java.util.List<String> qwords = new ArrayList<>();
                for (String q : qwordsArr) { if (!q.isEmpty()) qwords.add(q); }
                if (qwords.isEmpty()) return new MatchResult(0, -1, -1);
                java.util.List<Token> tokens = tokenize(lower);
                int count = 0; int firstStart = -1; int firstEnd = -1;
                int n = tokens.size();
                for (int i = 0; i < n; i++) {
                    if (!tokens.get(i).word.equals(qwords.get(0))) continue;
                    int idx = i; int lastIdx = i; boolean ok = true;
                    for (int qi = 1; qi < qwords.size(); qi++) {
                        String target = qwords.get(qi);
                        boolean found = false;
                        int limit = options.proximityN;
                        int searched = 0;
                        int j = lastIdx + 1;
                        while (j < n && searched <= limit) {
                            if (tokens.get(j).word.equals(target)) { found = true; break; }
                            j++; searched++;
                        }
                        if (!found) { ok = false; break; }
                        lastIdx = j;
                    }
                    if (ok) {
                        count++;
                        int s = tokens.get(i).start;
                        int e = tokens.get(lastIdx).end;
                        if (firstStart < 0) { firstStart = s; firstEnd = e; }
                        // continue search after this match to avoid overlapping starting at same i
                        i = lastIdx; // advance i to end of this match
                    }
                }
                return new MatchResult(count, firstStart, firstEnd);
            } else if (options.exactPhrase && options.ignoreSpaces) {
                // Two-pointer search ignoring whitespace in haystack
                String pattern = termLower.replaceAll("\\s+", "");
                if (pattern.isEmpty()) return new MatchResult(0, -1, -1);
                int n = lower.length();
                int m = pattern.length();
                int i = 0; int count = 0; int firstStart = -1; int firstEnd = -1;
                while (i < n) {
                    int tempI = i; int j = 0; int start = -1;
                    while (tempI < n && j < m) {
                        char ch = lower.charAt(tempI);
                        if (Character.isWhitespace(ch)) { tempI++; continue; }
                        if (ch == pattern.charAt(j)) {
                            if (start < 0) start = tempI;
                            tempI++; j++;
                        } else {
                            break;
                        }
                    }
                    if (j == m) {
                        count++;
                        if (firstStart < 0) { firstStart = start; firstEnd = tempI; }
                        i = tempI; // non-overlapping
                    } else {
                        i++;
                    }
                }
                return new MatchResult(count, firstStart, firstEnd);
            } else {
                // Simple substring search
                int count = 0; int firstStart = -1; int firstEnd = -1;
                int idx = 0;
                while ((idx = lower.indexOf(termLower, idx)) >= 0) {
                    if (firstStart < 0) { firstStart = idx; firstEnd = idx + termLower.length(); }
                    count++;
                    idx += termLower.length();
                }
                return new MatchResult(count, firstStart, firstEnd);
            }
        }

        private static String makeSnippet(String text, int matchStart, int matchEnd, int maxLen) {
            if (text == null || text.isEmpty()) return "";
            if (matchStart < 0 || matchEnd < 0 || matchStart >= text.length()) {
                // fallback to beginning
                return text.substring(0, Math.min(maxLen, text.length())).replaceAll("\n+", " ");
            }
            int mid = (matchStart + matchEnd) / 2;
            int start = Math.max(0, mid - maxLen / 2);
            int end = Math.min(text.length(), start + maxLen);
            // ensure we cover the match
            if (matchStart < start) start = Math.max(0, matchStart - 10);
            if (matchEnd > end) end = Math.min(text.length(), matchEnd + 10);
            String snippet = text.substring(start, end).replaceAll("\n+", " ");
            if (start > 0) snippet = "…" + snippet;
            if (end < text.length()) snippet = snippet + "…";
            return snippet;
        }

        // legacy snippet helper kept for backward compatibility
        private static String makeSnippet(String text, String needleLower, int maxLen) {
            if (text == null || text.isEmpty()) return "";
            String lower = text.toLowerCase(Locale.ROOT);
            int idx = lower.indexOf(needleLower);
            if (idx < 0) return text.substring(0, Math.min(maxLen, text.length())).replaceAll("\n+", " ");
            int start = Math.max(0, idx - maxLen / 2);
            int end = Math.min(text.length(), idx + needleLower.length() + maxLen / 2);
            String snippet = text.substring(start, end).replaceAll("\n+", " ");
            if (start > 0) snippet = "…" + snippet;
            if (end < text.length()) snippet = snippet + "…";
            return snippet;
        }
    }

    private interface CompletionCallback {
        void onCompleted(Throwable error, boolean cancelled, long filesScanned, long matchesFound, long elapsedMillis);
    }

    private static class SearchResult {
        final Path file;
        final int matchCount;
        final String snippet;
        SearchResult(Path file, int matchCount, String snippet) {
            this.file = file; this.matchCount = matchCount; this.snippet = snippet;
        }
    }

    // ODT text extractor without external libs
    static class OdtTextExtractor {
        static String extractText(Path odtFile) throws Exception {
            try (ZipFile zip = new ZipFile(odtFile.toFile())) {
                ZipEntry entry = zip.getEntry("content.xml");
                if (entry == null) throw new FileNotFoundException("content.xml not found in ODT: " + odtFile);
                try (InputStream in = zip.getInputStream(entry)) {
                    // Parse XML and get textContent
                    DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
                    dbf.setNamespaceAware(true);
                    // Secure processing
                    try { dbf.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true); } catch (Throwable ignored) {}
                    try { dbf.setFeature("http://xml.org/sax/features/external-general-entities", false); } catch (Throwable ignored) {}
                    try { dbf.setFeature("http://xml.org/sax/features/external-parameter-entities", false); } catch (Throwable ignored) {}
                    try { dbf.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false); } catch (Throwable ignored) {}
                    DocumentBuilder builder = dbf.newDocumentBuilder();
                    InputSource is = new InputSource(new BufferedReader(new InputStreamReader(in, StandardCharsets.UTF_8)));
                    Document doc = builder.parse(is);
                    String text = doc.getDocumentElement().getTextContent();
                    return normalizeWhitespace(text);
                }
            }
        }

        private static String normalizeWhitespace(String s) {
            if (s == null) return "";
            // Compress multiple spaces/newlines to single spaces, but keep basic separation
            return s.replace('\r', ' ').replace('\n', ' ').replace('\t', ' ').replaceAll(" +", " ").trim();
        }
    }

    // DOCX text extractor without external libs
    static class DocxTextExtractor {
        static String extractText(Path docxFile) throws Exception {
            StringBuilder sb = new StringBuilder();
            try (ZipFile zip = new ZipFile(docxFile.toFile())) {
                // Main document
                appendXmlText(zip, "word/document.xml", sb);
                // Headers and footers (header1.xml, header2.xml, etc.)
                Enumeration<? extends ZipEntry> entries = zip.entries();
                while (entries.hasMoreElements()) {
                    ZipEntry ze = entries.nextElement();
                    String name = ze.getName();
                    if (name.startsWith("word/header") && name.endsWith(".xml")) {
                        appendXmlText(zip, name, sb);
                    } else if (name.startsWith("word/footer") && name.endsWith(".xml")) {
                        appendXmlText(zip, name, sb);
                    } else if (name.equals("word/footnotes.xml") || name.equals("word/endnotes.xml")) {
                        appendXmlText(zip, name, sb);
                    }
                }
            }
            return normalizeWhitespace(sb.toString());
        }

        private static void appendXmlText(ZipFile zip, String path, StringBuilder out) {
            try {
                ZipEntry entry = zip.getEntry(path);
                if (entry == null) return;
                try (InputStream in = zip.getInputStream(entry)) {
                    DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
                    dbf.setNamespaceAware(true);
                    try { dbf.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true); } catch (Throwable ignored) {}
                    try { dbf.setFeature("http://xml.org/sax/features/external-general-entities", false); } catch (Throwable ignored) {}
                    try { dbf.setFeature("http://xml.org/sax/features/external-parameter-entities", false); } catch (Throwable ignored) {}
                    try { dbf.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false); } catch (Throwable ignored) {}
                    DocumentBuilder builder = dbf.newDocumentBuilder();
                    InputSource is = new InputSource(new BufferedReader(new InputStreamReader(in, StandardCharsets.UTF_8)));
                    Document doc = builder.parse(is);
                    String text = doc.getDocumentElement().getTextContent();
                    if (text != null && !text.isEmpty()) {
                        if (out.length() > 0) out.append('\n');
                        out.append(text);
                    }
                }
            } catch (Exception e) {
                // Ignore specific part errors; caller will log general failures per file
            }
        }

        private static String normalizeWhitespace(String s) {
            if (s == null) return "";
            return s.replace('\r', ' ').replace('\n', ' ').replace('\t', ' ').replaceAll(" +", " ").trim();
        }
    }

    // XLSX text extractor without external libs
    static class XlsxTextExtractor {
        static String extractText(Path xlsxFile) throws Exception {
            try (ZipFile zip = new ZipFile(xlsxFile.toFile())) {
                List<String> shared = readSharedStrings(zip);
                StringBuilder sb = new StringBuilder();
                Enumeration<? extends ZipEntry> entries = zip.entries();
                while (entries.hasMoreElements()) {
                    ZipEntry ze = entries.nextElement();
                    String name = ze.getName();
                    if (name.startsWith("xl/worksheets/") && name.endsWith(".xml")) {
                        String sheetText = readSheet(zip, name, shared);
                        if (!sheetText.isEmpty()) {
                            if (sb.length() > 0) sb.append('\n');
                            sb.append(sheetText);
                        }
                    }
                }
                return normalizeWhitespace(sb.toString());
            }
        }

        private static List<String> readSharedStrings(ZipFile zip) {
            ZipEntry entry = zip.getEntry("xl/sharedStrings.xml");
            if (entry == null) return Collections.emptyList();
            try (InputStream in = zip.getInputStream(entry)) {
                DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
                dbf.setNamespaceAware(true);
                try { dbf.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true); } catch (Throwable ignored) {}
                try { dbf.setFeature("http://xml.org/sax/features/external-general-entities", false); } catch (Throwable ignored) {}
                try { dbf.setFeature("http://xml.org/sax/features/external-parameter-entities", false); } catch (Throwable ignored) {}
                try { dbf.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false); } catch (Throwable ignored) {}
                DocumentBuilder builder = dbf.newDocumentBuilder();
                InputSource is = new InputSource(new BufferedReader(new InputStreamReader(in, StandardCharsets.UTF_8)));
                Document doc = builder.parse(is);
                NodeList siList = doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "si");
                if (siList.getLength() == 0) {
                    // Fallback without namespace
                    siList = doc.getElementsByTagName("si");
                }
                List<String> res = new ArrayList<>(siList.getLength());
                for (int i = 0; i < siList.getLength(); i++) {
                    Element si = (Element) siList.item(i);
                    String text = si.getTextContent();
                    res.add(text == null ? "" : text);
                }
                return res;
            } catch (Exception e) {
                return Collections.emptyList();
            }
        }

        private static String readSheet(ZipFile zip, String path, List<String> shared) {
            try (InputStream in = zip.getInputStream(zip.getEntry(path))) {
                DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
                dbf.setNamespaceAware(true);
                try { dbf.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true); } catch (Throwable ignored) {}
                try { dbf.setFeature("http://xml.org/sax/features/external-general-entities", false); } catch (Throwable ignored) {}
                try { dbf.setFeature("http://xml.org/sax/features/external-parameter-entities", false); } catch (Throwable ignored) {}
                try { dbf.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false); } catch (Throwable ignored) {}
                DocumentBuilder builder = dbf.newDocumentBuilder();
                InputSource is = new InputSource(new BufferedReader(new InputStreamReader(in, StandardCharsets.UTF_8)));
                Document doc = builder.parse(is);
                StringBuilder out = new StringBuilder();

                NodeList rowNodes = doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "row");
                if (rowNodes.getLength() == 0) rowNodes = doc.getElementsByTagName("row");
                for (int i = 0; i < rowNodes.getLength(); i++) {
                    Element row = (Element) rowNodes.item(i);
                    NodeList cellNodes = row.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "c");
                    if (cellNodes.getLength() == 0) cellNodes = row.getElementsByTagName("c");
                    for (int j = 0; j < cellNodes.getLength(); j++) {
                        Element c = (Element) cellNodes.item(j);
                        String t = c.getAttribute("t");
                        String cellText = "";
                        if ("s".equals(t)) {
                            // shared string
                            String idxStr = getFirstChildText(c, "v");
                            if (idxStr != null && !idxStr.isEmpty()) {
                                try {
                                    int idx = Integer.parseInt(idxStr.trim());
                                    if (idx >= 0 && idx < shared.size()) cellText = shared.get(idx);
                                } catch (NumberFormatException ignored) {}
                            }
                        } else if ("str".equals(t)) {
                            cellText = getFirstChildText(c, "v");
                        } else if ("inlineStr".equals(t)) {
                            // inline string stored under is/t
                            NodeList isNodes = c.getElementsByTagName("is");
                            if (isNodes.getLength() == 0) isNodes = c.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "is");
                            if (isNodes.getLength() > 0) {
                                Element isEl = (Element) isNodes.item(0);
                                cellText = isEl.getTextContent();
                            }
                        } else {
                            // numeric or date stored in v
                            cellText = getFirstChildText(c, "v");
                        }
                        if (cellText != null && !cellText.isEmpty()) {
                            if (out.length() > 0) out.append(' ');
                            out.append(cellText);
                        }
                    }
                    // new line per row
                    if (out.length() > 0) out.append('\n');
                }
                return out.toString();
            } catch (Exception e) {
                return "";
            }
        }

        private static String getFirstChildText(Element parent, String localName) {
            NodeList nodes = parent.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", localName);
            if (nodes.getLength() == 0) nodes = parent.getElementsByTagName(localName);
            if (nodes.getLength() == 0) return null;
            Node n = nodes.item(0);
            return n.getTextContent();
        }

        private static String normalizeWhitespace(String s) {
            if (s == null) return "";
            return s.replace('\r', ' ').replace('\n', ' ').replace('\t', ' ').replaceAll(" +", " ").trim();
        }
    }

    // Icon rendering factory (programmatic app icon with large "ODT" and small "search")
    static class IconFactory {
        static java.util.List<Image> createAppIcons() {
            int[] sizes = {16, 24, 32, 48, 64, 128, 256};
            java.util.List<Image> images = new ArrayList<>(sizes.length);
            for (int s : sizes) {
                images.add(renderIcon(s));
            }
            return images;
        }

        static BufferedImage renderIcon(int size) {
            BufferedImage img = new BufferedImage(size, size, BufferedImage.TYPE_INT_ARGB);
            Graphics2D g = img.createGraphics();
            try {
                g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

                // Background: rounded rectangle with vertical gradient
                float r = Math.max(4f, size * 0.18f);
                Color cTop = new Color(0x1E, 0x88, 0xE5);     // blue 600
                Color cBot = new Color(0x15, 0x65, 0xC0);     // blue 800
                Paint grad = new GradientPaint(0, 0, cTop, 0, size, cBot);
                g.setPaint(grad);
                g.fillRoundRect(0, 0, size, size, Math.round(r), Math.round(r));

                // Subtle inner highlight
                g.setColor(new Color(255, 255, 255, 28));
                int inset = Math.max(1, Math.round(size * 0.04f));
                g.fillRoundRect(inset, inset, size - inset * 2, Math.round(size * 0.46f), Math.round(r * 0.8f), Math.round(r * 0.8f));

                // Border
                g.setStroke(new BasicStroke(Math.max(1f, size * 0.03f)));
                g.setColor(new Color(0, 0, 0, 40));
                g.drawRoundRect(Math.round(0.5f), Math.round(0.5f), size - 1, size - 1, Math.round(r), Math.round(r));

                // Text settings
                String topText = "ODT";
                String bottomText = "search";

                // Compute font sizes relative to icon size
                float topFontSize = size * 0.50f;     // big letters
                float botFontSize = size * 0.18f;     // small label

                // Fit top text within width with some padding
                int pad = Math.round(size * 0.10f);
                Font topFont = fitText(g, topText, new Font("SansSerif", Font.BOLD, Math.round(topFontSize)), size - pad * 2);
                g.setFont(topFont);
                FontMetrics fmTop = g.getFontMetrics();
                int topTextWidth = fmTop.stringWidth(topText);
                int topTextAscent = fmTop.getAscent();
                int topX = (size - topTextWidth) / 2;
                int topY = Math.round(size * 0.58f); // vertically centered slightly above middle to make room for label

                // Draw slight shadow for readability
                g.setColor(new Color(0, 0, 0, 90));
                g.drawString(topText, topX + Math.max(1, size / 64), topY + Math.max(1, size / 64));
                g.setColor(Color.WHITE);
                g.drawString(topText, topX, topY);

                // Bottom small label on a translucent pill
                Font bottomFont = new Font("SansSerif", Font.PLAIN, Math.round(botFontSize));
                bottomFont = fitText(g, bottomText, bottomFont, size - pad * 2);
                g.setFont(bottomFont);
                FontMetrics fmBot = g.getFontMetrics();
                int botW = fmBot.stringWidth(bottomText);
                int botH = fmBot.getAscent();
                int pillPadX = Math.max(4, Math.round(size * 0.06f));
                int pillPadY = Math.max(2, Math.round(size * 0.03f));
                int pillW = botW + pillPadX * 2;
                int pillH = botH + pillPadY * 2;
                int pillX = (size - pillW) / 2;
                int pillY = size - pillH - Math.max(3, Math.round(size * 0.06f));

                // Pill background
                g.setColor(new Color(255, 255, 255, 140));
                g.fillRoundRect(pillX, pillY, pillW, pillH, Math.round(pillH * 0.6f), Math.round(pillH * 0.6f));
                g.setColor(new Color(0, 0, 0, 50));
                g.drawRoundRect(pillX, pillY, pillW, pillH, Math.round(pillH * 0.6f), Math.round(pillH * 0.6f));

                // Label text (dark)
                int textX = pillX + (pillW - botW) / 2;
                int textY = pillY + pillPadY + botH - Math.max(1, Math.round(size * 0.01f));
                g.setColor(new Color(25, 33, 44));
                g.drawString(bottomText, textX, textY);
            } finally {
                g.dispose();
            }
            return img;
        }

        private static Font fitText(Graphics2D g, String text, Font base, int maxWidth) {
            // Adjust font size down until it fits within maxWidth
            float size = base.getSize2D();
            if (size < 1f) size = 1f;
            Font f = base.deriveFont(size);
            FontMetrics fm = g.getFontMetrics(f);
            int w = fm.stringWidth(text);
            while (w > maxWidth && size > 1f) {
                size -= Math.max(1f, size * 0.06f);
                f = f.deriveFont(size);
                fm = g.getFontMetrics(f);
                w = fm.stringWidth(text);
            }
            return f;
        }

        // Minimal ICO writer that embeds PNG-compressed images (supported on Windows Vista+)
        static class IcoWriter {
            static void writePngIco(java.util.List<BufferedImage> images, OutputStream out) throws IOException {
                // Encode images to PNG bytes and sort by size ascending
                java.util.List<byte[]> pngs = new ArrayList<>(images.size());
                java.util.List<Integer> widths = new ArrayList<>(images.size());
                java.util.List<Integer> heights = new ArrayList<>(images.size());
                for (BufferedImage bi : images) {
                    ByteArrayOutputStream baos = new ByteArrayOutputStream();
                    ImageIO.write(bi, "png", baos);
                    byte[] data = baos.toByteArray();
                    pngs.add(data);
                    widths.add(bi.getWidth());
                    heights.add(bi.getHeight());
                }
                DataOutputStream dos = new DataOutputStream(out);
                // ICONDIR header
                dos.writeShort(0); // reserved
                dos.writeShort(1); // type: icon
                dos.writeShort(pngs.size()); // count
                int offset = 6 + (16 * pngs.size());
                for (int i = 0; i < pngs.size(); i++) {
                    int w = widths.get(i);
                    int h = heights.get(i);
                    byte bw = (byte) (w >= 256 ? 0 : w);
                    byte bh = (byte) (h >= 256 ? 0 : h);
                    dos.writeByte(bw); // width
                    dos.writeByte(bh); // height
                    dos.writeByte(0); // color count
                    dos.writeByte(0); // reserved
                    dos.writeShort(1); // planes
                    dos.writeShort(32); // bit count
                    dos.writeInt(pngs.get(i).length); // bytes in res
                    dos.writeInt(offset); // image offset
                    offset += pngs.get(i).length;
                }
                // image data blocks
                for (byte[] data : pngs) {
                    dos.write(data);
                }
                dos.flush();
            }
        }
    }

    // Log window
    static class LogWindow extends JFrame {
        private final JTextArea logArea = new JTextArea();
        LogWindow() {
            super("Search Log");
            setSize(800, 400);
            setLocationByPlatform(true);
            logArea.setEditable(false);
            logArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
            JScrollPane sp = new JScrollPane(logArea);
            getContentPane().add(sp);
        }
        void log(String message) {
            SwingUtilities.invokeLater(() -> {
                logArea.append(String.format("[%tF %<tT] %s%n", new Date(), message));
                logArea.setCaretPosition(logArea.getDocument().getLength());
            });
        }
    }

    public static void main(String[] args) {
        // Support headless icon export for packaging: --export-icon <path>
        if (args != null && args.length >= 2 && "--export-icon".equals(args[0])) {
            String outPath = args[1];
            try {
                java.util.List<Image> imgs = IconFactory.createAppIcons();
                java.util.List<BufferedImage> bi = new ArrayList<>(imgs.size());
                for (Image im : imgs) {
                    if (im instanceof BufferedImage) {
                        bi.add((BufferedImage) im);
                    } else {
                        BufferedImage buf = new BufferedImage(im.getWidth(null), im.getHeight(null), BufferedImage.TYPE_INT_ARGB);
                        Graphics2D g2 = buf.createGraphics();
                        try { g2.drawImage(im, 0, 0, null); } finally { g2.dispose(); }
                        bi.add(buf);
                    }
                }
                try (OutputStream os = Files.newOutputStream(Paths.get(outPath))) {
                    IconFactory.IcoWriter.writePngIco(bi, os);
                }
                System.out.println("[INFO] Icon exported to " + outPath);
                System.exit(0);
            } catch (Exception ex) {
                System.err.println("[ERROR] Failed to export icon: " + ex.getMessage());
                ex.printStackTrace();
                System.exit(2);
            }
            return;
        }
        SwingUtilities.invokeLater(() -> {
            try { UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); } catch (Exception ignored) {}
            new OdtSearchApp().setVisible(true);
        });
    }

    // ===== Selection persistence =====
    private void restoreLastSelectedDirectory() {
        String lastSel = null;
        try { lastSel = prefs.get(PREF_LAST_SELECTED, null); } catch (Exception ignore) {}
        DefaultMutableTreeNode rootNode = (DefaultMutableTreeNode) fileTree.getModel().getRoot();
        if (rootNode == null) return;
        // Default to root selection
        Runnable selectRoot = () -> {
            TreePath tp = new TreePath(rootNode.getPath());
            fileTree.setSelectionPath(tp);
            fileTree.scrollPathToVisible(tp);
        };
        if (lastSel == null || lastSel.isEmpty()) {
            selectRoot.run();
            return;
        }
        Path target;
        try {
            target = Paths.get(lastSel);
        } catch (Exception ex) {
            selectRoot.run();
            return;
        }
        try {
            if (!Files.isDirectory(target)) { selectRoot.run(); return; }
            Path absRoot = rootPath.toAbsolutePath().normalize();
            Path absTarget = target.toAbsolutePath().normalize();
            if (!absTarget.startsWith(absRoot)) { selectRoot.run(); return; }
            boolean ok = selectPathInTree(absTarget);
            if (!ok) selectRoot.run();
        } catch (Exception ex) {
            selectRoot.run();
        }
    }

    private boolean selectPathInTree(Path absTarget) {
        try {
            DefaultTreeModel model = (DefaultTreeModel) fileTree.getModel();
            DefaultMutableTreeNode node = (DefaultMutableTreeNode) model.getRoot();
            if (node == null) return false;
            FileNode fnRoot = (FileNode) node.getUserObject();
            Path absRoot = fnRoot.path.toAbsolutePath().normalize();
            if (absTarget.equals(absRoot)) {
                TreePath tp = new TreePath(node.getPath());
                fileTree.setSelectionPath(tp);
                fileTree.scrollPathToVisible(tp);
                return true;
            }
            Path rel = absRoot.relativize(absTarget);
            for (Path name : rel) {
                // Ensure children are loaded for current node
                ensureChildrenLoaded(node, model);
                DefaultMutableTreeNode next = null;
                for (int i = 0; i < node.getChildCount(); i++) {
                    DefaultMutableTreeNode ch = (DefaultMutableTreeNode) node.getChildAt(i);
                    Object uo = ch.getUserObject();
                    if (uo instanceof FileNode) {
                        Path p = ((FileNode) uo).path.getFileName();
                        if (p != null && p.toString().equalsIgnoreCase(name.toString())) {
                            next = ch; break;
                        }
                    }
                }
                if (next == null) return false;
                node = next;
                // Expand path as we go
                TreePath tp = new TreePath(node.getPath());
                fileTree.expandPath(tp);
            }
            TreePath tp = new TreePath(node.getPath());
            fileTree.setSelectionPath(tp);
            fileTree.scrollPathToVisible(tp);
            return true;
        } catch (Exception ex) {
            return false;
        }
    }

    private void ensureChildrenLoaded(DefaultMutableTreeNode node, DefaultTreeModel model) {
        if (node.getChildCount() == 1 && ((DefaultMutableTreeNode) node.getFirstChild()).getUserObject() instanceof String) {
            // Placeholder present: replace with real children
            node.removeAllChildren();
            Object uo = node.getUserObject();
            if (uo instanceof FileNode) {
                loadChildren(node);
            }
            model.reload(node);
        } else if (node.getChildCount() == 0) {
            Object uo = node.getUserObject();
            if (uo instanceof FileNode) {
                loadChildren(node);
                model.reload(node);
            }
        }
    }

}