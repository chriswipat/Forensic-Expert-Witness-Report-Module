/*
 * Class ForensicReportConfigPanel.java of package ForensicReport
 * 
 * Using this class you are able to display a graphical user interface (GUI) in the report
 * generating section of Autopsy, and allow the user to select the files which he or she
 * would like to include. This GUI allows the user to select which forensic expert 
 * witness report which he or she would like to report into and also brings the 
 * user the option to select from three included forensic expert witness 
 * report templates.
 * 
 * This Report Module was written for a final year project for the degree
 * of Computer and Digital Forensics BSc (Hons) at Northumbria University 
 * in Newcastle and improved for OSDFCon, 2018. The final year project at
 * Northumbria University included the aim of aiding in the automation, 
 * ease and effectiveness of digital forensic practitioners whilst 
 * conducting digital forensic investigations in Autopsy.
 * 
 * @author Chris Wipat
 * @version 17.09.2018
 */

package ForensicExpertWitnessReport;

import java.awt.Component;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import javax.swing.JCheckBox;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.ListCellRenderer;
import javax.swing.ListModel;
import javax.swing.event.ListDataListener;
import org.sleuthkit.autopsy.casemodule.Case;
import org.sleuthkit.autopsy.coreutils.Logger;
import org.sleuthkit.datamodel.TagName;
import org.sleuthkit.datamodel.TskCoreException;
import javax.swing.JFileChooser;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

class ForensicReportConfigPanel extends javax.swing.JPanel {

    // Declare Instance Variables
    private final Map<String, Boolean> tagNameSelections = new LinkedHashMap<String, Boolean>();
    private final TagNamesListModel tagsNamesListModel = new TagNamesListModel();  
    private final TagsNamesListCellRenderer tagsNamesRenderer = new TagsNamesListCellRenderer();
    private List<TagName> tagNames;
    private static final long serialVersionUID = 1L; 
    private final Set<String> supported_extentions = new HashSet<String>();
    private final String TemplateOne_name = "Pre-existing Template 1";
    private final String TemplateTwo_name = "Pre-existing Template 2"; 
    private final String TemplateThree_name = "Pre-existing Template 3";
    private XWPFDocument TemplateOne_doc = null;
    private XWPFDocument TemplateTwo_doc = null;
    private XWPFDocument TemplateThree_doc = null; 
//    private XWPFDocument Example_table = null;
    private XWPFDocument inputted_doc = null;
    private XWPFDocument selected_doc = null;
    private boolean TemplateTwo_extracted = false;
    private boolean TemplateThree_extracted = false;
    private String inputted_name = "input";
    private String inputted_full_path;
    private String inputted_file_ext;       
    private String selectedDocumentName = TemplateOne_name;
    private String evidenceHeading = "Analysis Evidence";
    private FileOutputStream fos = null;
    private File file = null;
    private File Dir = null;
    private String colourName = "";
    private String hexadecimalColourCode = "";
    private ArrayList<String> colourNames = new ArrayList<String>();
    private ArrayList<String> hexadecimalColourCodes = new ArrayList<String>();
    
    /**
     * Constructor for objects of class ForensicReportConfigPanel
     * First and only Constructor.
     * 
     * Call methods which populate GUI components and display the GUI to the user,
     * extract the documents from the compiled JAR after NBM installation, and the method
     * to create the document objects.
     * 
     * Includes Tag Name List Box, Forensic Expert Witness Report ComboBox & File Selector button. 
     */
    ForensicReportConfigPanel() {        
        initComponents();
        populateTagNameComponents();
        populateForensicReports();
        populateSupportedExtentions();
        extractDocument("Pre_existing_template_one.docx");
//        extractDocument("Example_Table.docx");
        createDocuments(null);
    }
        
    /**
     * PopulateTagNameComponents method
     * First Mutator Method.
     * 
     * Populates the Tag Name Components to the current tags which are 
     * in use for the current case in Autopsy, as selected and created by
     * the user. Set all of the tag name components in the list box to unselected
     * until the user selects the tags of files in which he wants to include.
     * 
     */
    private void populateTagNameComponents() {
        
        // Get the tag names in use for the current case, using imported Case class.
        try {
            tagNames = Case.getCurrentCase().getServices().getTagsManager().getTagNamesInUse();
        } catch (TskCoreException ex) {
            Logger.getLogger(ForensicReportConfigPanel.class.getName()).log(Level.SEVERE, "Failed to get tag names", ex);
            JOptionPane.showMessageDialog(null, "Error getting tag names for case.", "Tag Names Not Found", JOptionPane.ERROR_MESSAGE);
        }
        // Mark the tag names as unselected. Note that tagNameSelections is a
        // LinkedHashMap so that order is preserved and the tagNames and tagNameSelections
        // containers are "parallel" containers.
        for (TagName tagName : tagNames) {
            tagNameSelections.put(tagName.getDisplayName(), Boolean.FALSE);
        }
        // Set up the tag names JList component to be a collection of check boxes
        // for selecting tag names. The mouse click listener updates tagNameSelections
        // to reflect user choices.
        tagNamesListBox.setModel(tagsNamesListModel);
        tagNamesListBox.setCellRenderer(tagsNamesRenderer);
        tagNamesListBox.setVisibleRowCount(-1);
        tagNamesListBox.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent evt) {
                JList<?> list = (JList) evt.getSource();
                int index = list.locationToIndex(evt.getPoint());
                if (index > -1) {
                    String value = tagsNamesListModel.getElementAt(index);
                    tagNameSelections.put(value, !tagNameSelections.get(value));
                    list.repaint();
                }
            }
        });
    }
    
    /**
     * PopulateForensicReports method
     * Second Mutator Method.
     * 
     * Populate the ComboBox with the names of Forensic Expert Witness Reports,
     * as have been declared earlier in the instance variables.
     * 
     */
    private void populateForensicReports() {       
        expertWitnessReportComboBox.removeAllItems();        
        expertWitnessReportComboBox.addItem(TemplateOne_name); 
        expertWitnessReportComboBox.addItem(TemplateTwo_name);
        expertWitnessReportComboBox.addItem(TemplateThree_name);
    } 
    
    /**
     * InitComponents method
     * Third Mutator Method.
     * 
     * Set the GUI of every component and display the GUI to the user.
     * 
     * Created using NetBeans GUI designer.
     * 
     * Includes the following:
     * 
     * jLabel
     * jScrollPane
     * tagNamesListBox
     * selectAllButton
     * deselectAllButton 
     * jLabel2
     * expertWitnessReportComboBox   
     * chooseExpertWitnessReportButton
     * jLabel3
     * jTextField1
     */
    private void initComponents() {

    // Initialize declared instance variables
    jLabel1 = new javax.swing.JLabel();
    jScrollPane1 = new javax.swing.JScrollPane();
    tagNamesListBox = new javax.swing.JList<String>();   
    selectAllButton = new javax.swing.JButton();
    deselectAllButton = new javax.swing.JButton();
    optionsButton = new javax.swing.JButton();
    jLabel2 = new javax.swing.JLabel();
    expertWitnessReportComboBox = new javax.swing.JComboBox<String>();
    chooseExpertWitnessReportButton = new javax.swing.JButton();
    jLabel3 = new javax.swing.JLabel();
    jTextField1 = new javax.swing.JTextField();    

   org.openide.awt.Mnemonics.setLocalizedText(jLabel1, "Export files tagged as:");
    
    jScrollPane1.setViewportView(tagNamesListBox);
     
    org.openide.awt.Mnemonics.setLocalizedText(selectAllButton, "Select all");
    selectAllButton.addActionListener(new java.awt.event.ActionListener() {
        public void actionPerformed(java.awt.event.ActionEvent evt) {
            selectAllButtonActionPerformed(evt);
        }
    });
 
    org.openide.awt.Mnemonics.setLocalizedText(deselectAllButton, "Deslect all");
    deselectAllButton.addActionListener(new java.awt.event.ActionListener() {
        public void actionPerformed(java.awt.event.ActionEvent evt) {
            deselectAllButtonActionPerformed(evt);
        }
    });
    
    org.openide.awt.Mnemonics.setLocalizedText(optionsButton, "Options");
    optionsButton.addActionListener(new java.awt.event.ActionListener() {
        public void actionPerformed(java.awt.event.ActionEvent evt) {
            optionsButtonActionPerformed(evt);
        }
    });
    
    org.openide.awt.Mnemonics.setLocalizedText(jLabel2, "Select Forensic Expert Witness Report");
    
    expertWitnessReportComboBox.addActionListener(new java.awt.event.ActionListener() {
        public void actionPerformed(java.awt.event.ActionEvent evt) {
            expertWitnessReportComboBoxActionPerformed(evt);
        }
    });
 
    org.openide.awt.Mnemonics.setLocalizedText(chooseExpertWitnessReportButton, "Choose file");
    chooseExpertWitnessReportButton.addActionListener(new java.awt.event.ActionListener() {
        public void actionPerformed(java.awt.event.ActionEvent evt) {
            chooseExpertWitnessReportButtonActionPerformed(evt);
        }
    });
    
    org.openide.awt.Mnemonics.setLocalizedText(jLabel3, "Enter Heading or Sub-heading:");
    
    jTextField1.setText(evidenceHeading);
    jTextField1.addKeyListener(new java.awt.event.KeyAdapter() {
        @Override
        public void keyReleased(java.awt.event.KeyEvent evt) {
            jTextField1KeyReleased(evt);
        }
    });
 
    javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addComponent(jLabel2)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jTextField1, javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel3, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 323, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    .addComponent(selectAllButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(deselectAllButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(optionsButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                            .addComponent(jLabel1)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(chooseExpertWitnessReportButton)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(expertWitnessReportComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, 162, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel1)
                .addGap(8, 8, 8)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(selectAllButton)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(deselectAllButton)
                        .addGap(18, 18, 18)
                        .addComponent(optionsButton))
                    .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 106, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel3)
                    .addComponent(jLabel2))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(expertWitnessReportComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(chooseExpertWitnessReportButton)))
                .addContainerGap())
        );
    }
    
    /**
     * SelectAllButtonActionPerformed method
     * Fourth Mutator Method.
     * 
     * On button pressed, for each tag name in the case, set it to selected.
     * 
     * @param evt 
     */
    private void selectAllButtonActionPerformed(java.awt.event.ActionEvent evt) {
        for (TagName tagName : tagNames) {
            tagNameSelections.put(tagName.getDisplayName(), Boolean.TRUE);
        }
        tagNamesListBox.repaint();
    }//GEN-LAST:event_selectAllButtonActionPerformed 
 
    /**
     * DeselectAllButtonActionPerformed Method
     * Fifth Mutator Method.
     * 
     * On button pressed, for each tag name in the case, set it to un-selected.
     * 
     * @param evt 
     */
    private void deselectAllButtonActionPerformed(java.awt.event.ActionEvent evt) {
        for (TagName tagName : tagNames) {
            tagNameSelections.put(tagName.getDisplayName(), Boolean.FALSE);
        }
        tagNamesListBox.repaint();
    }//GEN-LAST:event_deselectAllButtonActionPerformed
    
    /**
     * OptionsButtonActionPerformed Method
     * Sixth Mutator Method.
     * 
     * On button pressed, load the options panel.
     * 
     * @param evt 
     */
    private void optionsButtonActionPerformed(java.awt.event.ActionEvent evt) {
        ForensicReportConfigPanelOptions configPanel = new ForensicReportConfigPanelOptions(colourName, hexadecimalColourCode);
        configPanel.load();
        if (JOptionPane.showConfirmDialog(null, configPanel, "Forensic Expert Witness Report Options", JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE) == JOptionPane.OK_OPTION) {
             configPanel.store();
             colourName = configPanel.getTableColourName();
             hexadecimalColourCode = configPanel.getHexadecimalColourCode();
             colourNames = configPanel.getColourNames();
             hexadecimalColourCodes = configPanel.getColourHexadecimals();
             
            if (hexadecimalColourCode.trim().length() != 6) {               
                JOptionPane.showMessageDialog(null, "Hexadecimal colour codes must be 6 characters.", "Incorrect hexadecimal code.", JOptionPane.ERROR_MESSAGE);
            }

        }
    }
    
    /**
     * ExpertWitnessReportComboBoxActionPerformed Method
     * Seventh Mutator Method.
     * 
     * On Combo Box user selection, set the selected item to an instance variable,
     * update the evidence heading or sub-heading text field with the matching
     * forensic expert witness report heading or sub-heading.
     * 
     * @param evt 
     */
    private void expertWitnessReportComboBoxActionPerformed(java.awt.event.ActionEvent evt) {
        
        selectedDocumentName = (String)expertWitnessReportComboBox.getSelectedItem();
        if (TemplateOne_name.equals(selectedDocumentName)) {
            jTextField1.setText("Analysis Evidence");
            evidenceHeading = "Analysis Evidence";
            selected_doc = TemplateOne_doc;
        }
        if (TemplateTwo_name.equals(selectedDocumentName)) {
            if(!TemplateTwo_extracted) {
                extractDocument("Pre_existing_template_two.docx");
                TemplateTwo_extracted = true;
                createDocuments(null);
                selected_doc = TemplateTwo_doc;
            }
            jTextField1.setText("Analysis Evidence");
            evidenceHeading = "Analysis Evidence";
        }
        if (TemplateThree_name.equals(selectedDocumentName)) {
            if(!TemplateThree_extracted) {                
                extractDocument("Pre_existing_template_three.docx");
                TemplateThree_extracted = true;
                createDocuments(null);
                selected_doc = TemplateThree_doc;
            }
            jTextField1.setText("Section 2 - Evidence");
            evidenceHeading = "Section 2 - Evidence";
        }
        if (selectedDocumentName != null && inputted_name != null) {
            if (selectedDocumentName.equals(inputted_name)) {                
                jTextField1.setText("Enter a heading");
                evidenceHeading = "";
                selected_doc = inputted_doc;
            }
        }
    
    }//GEN-LAST:event_hashSetsComboBoxActionPerformed

    /**
     * ChooseExpertWitnessReportButtonActionPerformed Method
     * Eighth Mutator Method.
     * 
     * On Choose File button selected, declare JFileChooser and show
     * the file chooser to the user. Set the selected files and it's
     * full path to 2 instance variables, re-populate the ComboBox with 
     * the chosen file, and re-create the document objects for the 
     * forensic expert witness reports. 
     * 
     * @param evt 
     */
    private void chooseExpertWitnessReportButtonActionPerformed(java.awt.event.ActionEvent evt) {
    
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("Microsoft Word Documents", "DOCX", "DOTX", "DOTM", "DOCM", "DOCM FlatOPC")); 
        fileChooser.setAcceptAllFileFilterUsed(true);
        int result = fileChooser.showOpenDialog(null);
 
        if (result == JFileChooser.APPROVE_OPTION) {            
            inputted_name = fileChooser.getSelectedFile().getName();
            inputted_full_path = fileChooser.getSelectedFile().getAbsolutePath();
            inputted_file_ext = FilenameUtils.getExtension(inputted_name);
            if (supported_extentions.contains(inputted_file_ext.toLowerCase()))
            {
                populateForensicReports();
                createDocuments(inputted_full_path);
                expertWitnessReportComboBox.addItem(inputted_name); 
                expertWitnessReportComboBox.setSelectedIndex(3);
            }            
        }    
        
    }
    
    /**
     * JTextField1KeyReleased Method
     * Ninth Mutator Method.
     * 
     * On key release, set the inputted text to an instance variable.
     * 
     * @param evt 
     */
     private void jTextField1KeyReleased(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_jTextField1KeyReleased
         evidenceHeading = jTextField1.getText();
     }//GEN-LAST:event_jTextField1KeyReleased
    
    /**
     * CreateDocuments
     * Tenth Mutator Method.
     * 
     * Creates Document Objects for Forensic Expert Witness Reports
     * 
     * @param inputted 
     */
    private void createDocuments(String inputdoc) {
        try {
            TemplateOne_doc = new XWPFDocument(new FileInputStream(System.getProperty("user.home") + "\\.ForensicReportModule\\Pre_existing_template_one.docx"));
//            Example_table = new XWPFDocument(new FileInputStream(System.getProperty("user.home") + "\\.ForensicReportModule\\Example_Table.docx"));
            if (TemplateTwo_extracted) {
                TemplateTwo_doc = new XWPFDocument(new FileInputStream(System.getProperty("user.home") + "\\.ForensicReportModule\\Pre_existing_template_two.docx"));
            }
            if (TemplateThree_extracted) {
                TemplateThree_doc = new XWPFDocument(new FileInputStream(System.getProperty("user.home") + "\\.ForensicReportModule\\Pre_existing_template_three.docx"));
            }
            if (inputdoc != null && !inputdoc.isEmpty()) {
                inputted_doc = new XWPFDocument(new FileInputStream(inputdoc));
            }
        }
        catch(IOException e){
            Logger.getLogger(ForensicReportConfigPanel.class.getName()).log(Level.SEVERE, "Failed to create document objects", e);
            JOptionPane.showMessageDialog(null, "Document object error", "Failed to create document objects", JOptionPane.ERROR_MESSAGE);
        }
    } 
    
    /**
     * ExtractDocument
     * Eleventh Mutator Method.
     * 
     * Extracts Pre-Existing Templates from NetBeans/JAR Package into the User Home Directory.
     * 
     * @param inputted 
     */
    private void extractDocument(String document) {
        try {
            // Declare InputStream object as the document from the Java package / compiled JAR
            InputStream in = getClass().getResourceAsStream(document);
            
            // Create new file object as Dir, set to user home / .ForensicReportModule Directory
            String dir = System.getProperty("user.home") + "\\.ForensicReportModule";
            Dir = new File(dir);
            
            // If directory doesn't exist, create it
            if (!Dir.exists()) {
                try{
                    Dir.mkdir();
                } 
                catch(SecurityException se){
                    Logger.getLogger(ForensicReportConfigPanel.class.getName()).log(Level.SEVERE, "Error creating folder " +Dir, se);
                }        
            }
            
            // Declare new file object, set to User home / .ForensicReportModule Directory + document           
            file = new File(dir + "\\" +document);
            
            // Create new FileOutputStream object as the created file object
            fos = new FileOutputStream(file);

            //If the file doesn't exist, create the file
	    if (!file.exists()) {
                file.createNewFile();
            }

            // Convert InputStream object / document to bytesArray
            byte[] bytesArray = IOUtils.toByteArray(in);

            // Write the bytesArray to the created file
            fos.write(bytesArray);
            fos.flush();            
        } 
        catch (IOException ioe) {
            Logger.getLogger(ForensicReportConfigPanel.class.getName()).log(Level.SEVERE, "Error extracting " +document + " from JAR package", ioe);
        } 
        finally {
            try {
                if (fos != null) 
                {
                    fos.close();
                }
            }
            catch (IOException ioe) {
                System.out.println("Error in closing the Stream");
            }
        }
    }
    
    /**
     * PopulateSupportedExtentions
     * Twelfth Mutator Method.
     * 
     * Add supported forensic expert witness report file extensions
     * 
     */
    private void populateSupportedExtentions () {
        supported_extentions.add("docx"); supported_extentions.add("dotx"); supported_extentions.add("dotm");
        supported_extentions.add("docm"); supported_extentions.add("docm flatopc"); 
    }
    
    /**
     * GetSelectedDocument Method
     * First Accessor Method.
     * 
     * Returns the document object for the selected forensic expert witness report.
     * 
     * @return inputted_doc
     * @return TemplateOne_doc
     * @return TemplateTwo_doc
     * @return TemplateThree_doc
     */
    public XWPFDocument getSelectedDocument() {
        if (selected_doc != null ) {
            return selected_doc;
        }
        return TemplateOne_doc;
   }
    
    /**
     * GetSelectedDocumentName Method
     * Second Accessor Method.
     * 
     * Returns the name of the selected document.
     * 
     * @return selectedDocumentName
     */
    public String getSelectedDocumentName() {
        return selectedDocumentName;
   }
    
    /**
     * Template_1_or_2_isSelected Method
     * Third Accessor Method
     * 
     * Returns whether template one or two is selected.
     * 
     * @return true
     * @return false
     */
    public boolean Template_1_or_2_isSelected() {
        return selectedDocumentName.equals(TemplateOne_name) || selectedDocumentName.equals(TemplateTwo_name);
    }
    
    /**
     * GetEvidenceHeading Method
     * Fourth Accessor Method.
     * 
     * Returns the name of the evidence heading or
     * sub-heading.
     * 
     * @return evidenceHeading
     */
    public String getEvidenceHeading() {
        return evidenceHeading;
    }
    
    /**
     * GetFileExtension Method
     * Fifth Accessor Method.
     * 
     * Returns the file extension of the selected document
     * 
     * @return file_extension
     */
    public String getFileExtension() {
        if (selectedDocumentName.equals(inputted_name)) {
            return inputted_file_ext;
        }
        if (selectedDocumentName.equals(TemplateOne_name)) {
            return "docx";
        }
        if (selectedDocumentName.equals(TemplateTwo_name)) {
            return "docx";
        }
        if (selectedDocumentName.equals(TemplateThree_name)) {
            return "docx";
        }
        return "docx";
    }
    
    /**
     * GetSelectedTagNames Method
     * Sixth Accessor Method.
     * 
     * Returns the user selected tag names for files he wishes to extract.
     * 
     * @return selectedTagNames
     */
    public List<TagName> getSelectedTagNames() {
        List<TagName> selectedTagNames = new ArrayList<TagName>();
        for (TagName tagName : tagNames) {
            if (tagNameSelections.get(tagName.getDisplayName())) {
                selectedTagNames.add(tagName);
            }
        }
        return selectedTagNames;
    }
    
     /**
     * ReturnTableColour Method
     * Seventh Accessor Method.
     * 
     * Return the selected table colour in hexadecimal to ForensicReport.java.
     * 
     * @return hexadecimalColourCodes.get(i)
     * @return hexadecimalColourCode
     * @return 000000
     */
    public String getTableColour() {
        
        // If hexadecimal value is bogus, use the distinguished colour.
        if (hexadecimalColourCode.trim().length() != 6) {
            
            // Match the corresponding hexadecimal value for the colour.
            for (int i=0; i<colourNames.size(); i++)
            {
                if(colourName.equals(colourNames.get(i))) {
                    return hexadecimalColourCodes.get(i);
                }
            }
            
        // If hexadecimal value isn't bogus, return it.
        } else {
            return hexadecimalColourCode;
        }
        
        return "000000";

    }
    
//    /**
//     * GetExampleTable Method.
//     * Seventh Accessor Method.
//     * 
//     * Return the document object for the example table.
//     * 
//     * Used later on to duplicate the spacing and style of the table.
//     * 
//     * @return Example_table
//     */
//    public XWPFDocument getExampleTable() {
//        return Example_table;
//    }
    
    /**
     * Class TagNamesListModel of package ForensicReport
     * 
     * Created in order to manage tag names provided by Autopsy in the
     * current case, and selected by the user in the GUI. Implements
     * ListModel.
     * 
     */
    private class TagNamesListModel implements ListModel<String> {

        @Override
        public int getSize() {
            return tagNames.size();
        }
	
	@Override
        public String getElementAt(int index) {
            return tagNames.get(index).getDisplayName();
        }
	
	@Override
        public void addListDataListener(ListDataListener l) {
        }
	
	@Override
        public void removeListDataListener(ListDataListener l) {
        }
    
    }
    
    /**
     * Class TagsNamesListCellRenderer of package ForensicReport
     * 
     * Created in order to render the items in the tag names JList component
     * as JCheckbox components. Extends JCheckBox, Implements ListCellRenderer.
     * 
     */
    private class TagsNamesListCellRenderer extends JCheckBox implements ListCellRenderer<String> {
        private static final long serialVersionUID = 1L;

        @Override
        public Component getListCellRendererComponent(JList<? extends String> list, String value, int index, boolean isSelected, boolean cellHasFocus) {
            if (value != null) {
                setEnabled(list.isEnabled());
                setSelected(tagNameSelections.get(value));
                setFont(list.getFont());
                setBackground(list.getBackground());
                setForeground(list.getForeground());
                setText(value);
                return this;
            }
            return new JLabel();
        }
    }      

    // GUI Variables declaration //GEN-BEGIN:variables    
    private javax.swing.JLabel jLabel1;
    private javax.swing.JScrollPane jScrollPane1;   
    private javax.swing.JList<String> tagNamesListBox = new JList<String>();
    private javax.swing.JButton optionsButton;
    private javax.swing.JButton selectAllButton;
    private javax.swing.JButton deselectAllButton;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JComboBox<String> expertWitnessReportComboBox;
    private javax.swing.JButton chooseExpertWitnessReportButton;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JLabel jLabel3;
    // End of variables declaration//GEN-END:variables    
    
}
