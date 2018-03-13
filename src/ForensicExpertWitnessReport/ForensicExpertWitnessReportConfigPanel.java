/*
 * Class ForensicExpertWitnessReportConfigPanel.java of package ForensicExpertwitnessReport
 * 
 * Using this class you are able to display a graphical user interface (GUI)
 * inside Autopsy which allows the user to select which tagged files he or
 * she would like to report, and the forensic expert witness report he or she
 * would like to report to. This further allows the selection of three included
 * forensic expert witness report templates.
 * 
 * This class was written for a final year project for
 * the degree of Computer and Digital Forensics BSc (Hons),
 * at Northumbria University in Newcastle. This project includes
 * the aim of aiding in automation, ease and effectivness of digital 
 * forensic practitioners while conducting digital forensic
 * investigations in Autopsy.
 * 
 * @author Chris Wipat
 * @version 19.04.2018
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
import com.aspose.words.Document;
import java.net.URL;
import java.util.ArrayList;

class ForensicExpertWitnessReportConfigPanel extends javax.swing.JPanel {

    // Declare Instance Variables
    private final Map<String, Boolean> tagNameSelections = new LinkedHashMap<String, Boolean>();
    private final TagNamesListModel tagsNamesListModel = new TagNamesListModel();  
    private final TagsNamesListCellRenderer tagsNamesRenderer = new TagsNamesListCellRenderer();
    private List<TagName> tagNames;
    private static final long serialVersionUID = 1L; 
    private final String TemplateOne_name = "Pre-existing Template 1";
    private final String TemplateTwo_name = "Pre-existing Template 2";
    private final String TemplateThree_name = "Pre-existing Template 3";   
    private Document TemplateOne_doc = null;
    private Document TemplateTwo_doc = null;
    private Document TemplateThree_doc = null;
    private String inputted_name = null;
    private String inputted_full_path;
    private Document inputted_doc = null;    
    private String selectedDocumentName = null;
    private String evidenceHeading = "Section 4. Evidence";
    
    /**
     * Constructor for objects of class ForensicExpertWitnessReportConfigPanel
     * First and only Constructor.
     * 
     * Call methods which populate GUI components and display the GUI to the user.
     * Call method to create objects for documents.
     * 
     * Includes Tag Name List Box, Forensic Expert Witness Report ComboBox & File Selector button. 
     */
    ForensicExpertWitnessReportConfigPanel() {
        initComponents();
        populateTagNameComponents();
        populateForensicExpertWitnessReports();
        createDocuments("");
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
            Logger.getLogger(ForensicExpertWitnessReportConfigPanel.class.getName()).log(Level.SEVERE, "Failed to get tag names", ex);
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
     * PopulateForensicExpertWitnessReports method
     * Second Mutator Method.
     * 
     * Populate the ComboBox with the names of Forensic Expert Witness Reports,
     * as have been declared earlier in the instance variables.
     * 
     */
    private void populateForensicExpertWitnessReports() {       
        expertWitnessReportComboBox.removeAllItems();        
        expertWitnessReportComboBox.addItem(TemplateOne_name); 
        expertWitnessReportComboBox.addItem(TemplateTwo_name);
        expertWitnessReportComboBox.addItem(TemplateThree_name);
    } 
    
    /**
     * InitComponents method
     * Third Mutator Method
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
    jLabel2 = new javax.swing.JLabel(); 
    expertWitnessReportComboBox = new javax.swing.JComboBox<String>();    
    chooseExpertWitnessReportButton = new javax.swing.JButton();
    jLabel3 = new javax.swing.JLabel();  
    jTextField1 = new javax.swing.JTextField();      

    org.openide.awt.Mnemonics.setLocalizedText(jLabel1, "Export files tagged as:"); // NOI18N 
    
    jScrollPane1.setViewportView(tagNamesListBox);
     
    org.openide.awt.Mnemonics.setLocalizedText(selectAllButton, "Select all"); // NOI18N
    selectAllButton.addActionListener(new java.awt.event.ActionListener() {
        public void actionPerformed(java.awt.event.ActionEvent evt) {
            selectAllButtonActionPerformed(evt);
        }
    });
 
    org.openide.awt.Mnemonics.setLocalizedText(deselectAllButton, "Deslect all"); // NOI18N
    deselectAllButton.addActionListener(new java.awt.event.ActionListener() {
        public void actionPerformed(java.awt.event.ActionEvent evt) {
            deselectAllButtonActionPerformed(evt);
        }
    }); 
    
    org.openide.awt.Mnemonics.setLocalizedText(jLabel2, "Select Forensic Expert Witness Report"); // NOI18N
    
    expertWitnessReportComboBox.addActionListener(new java.awt.event.ActionListener() {
        public void actionPerformed(java.awt.event.ActionEvent evt) {
            expertWitnessReportComboBoxActionPerformed(evt);
        }
    });
 
    org.openide.awt.Mnemonics.setLocalizedText(chooseExpertWitnessReportButton, "Choose file"); // NOI18N
    chooseExpertWitnessReportButton.addActionListener(new java.awt.event.ActionListener() {
        public void actionPerformed(java.awt.event.ActionEvent evt) {
            chooseExpertWitnessReportButtonActionPerformed(evt);
        }
    });
    
    org.openide.awt.Mnemonics.setLocalizedText(jLabel3, "Enter Heading or Sub-heading:"); // NOI18N
    
    jTextField1.setText(evidenceHeading); // NOI18N
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
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jLabel2)
                        .addGap(88, 88, 88)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jTextField1, javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel3, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 313, javax.swing.GroupLayout.DEFAULT_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(selectAllButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(deselectAllButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                    .addComponent(jLabel1)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(expertWitnessReportComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, 162, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(chooseExpertWitnessReportButton)))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel1)
                .addGap(8, 8, 8)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(selectAllButton)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(deselectAllButton))
                    .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 114, javax.swing.GroupLayout.DEFAULT_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel3)
                    .addComponent(jLabel2))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(chooseExpertWitnessReportButton)
                        .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(expertWitnessReportComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap())
        );
}
    
    /**
     * 
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
     * 
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
     * ExpertWitnessReportComboBoxActionPerformed Method
     * Sixth Mutator Method.
     * 
     * On Combo Box user selection, set the selected item to an instance variable,
     * update the evidence heading or sub-heading text field with the matching
     * forensic expert witness report heading or sub-heading.
     * 
     * @param evt 
     */
    private void expertWitnessReportComboBoxActionPerformed(java.awt.event.ActionEvent evt) {
        selectedDocumentName = (String)expertWitnessReportComboBox.getSelectedItem();
        if (null != selectedDocumentName) {
            if (TemplateOne_name.equals(selectedDocumentName)) {
                jTextField1.setText("Section 4. Evidence");
            }
            if (TemplateTwo_name.equals(selectedDocumentName)) {
                jTextField1.setText("Section 5. Evidence");
            }
            if (TemplateThree_name.equals(selectedDocumentName)) {
                jTextField1.setText("Section 3. Evidence");
            }
            if (selectedDocumentName.equals(inputted_name)) {
                jTextField1.setText("...");
            }
        }
        
    }//GEN-LAST:event_hashSetsComboBoxActionPerformed

    /**
     * ChooseExpertWitnessReportButtonActionPerformed Method
     * Seventh Mutator Method.
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
        fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);  
        int result = fileChooser.showOpenDialog(null);
 
        if (result == JFileChooser.APPROVE_OPTION) {            
            inputted_name = fileChooser.getSelectedFile().getName();
            inputted_full_path = fileChooser.getSelectedFile().getAbsolutePath();
            populateForensicExpertWitnessReports();
            createDocuments(inputted_full_path);
            expertWitnessReportComboBox.addItem(inputted_name); 
            expertWitnessReportComboBox.setSelectedIndex(3);
        }    
    }    
    /**
     * JTextField1KeyReleased Method
     * Eighth Mutator Method.
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
     * Ninth Mutator Method.
     * 
     * Creates Document Objects for Forensic Expert Witness Reports
     * 
     * @param inputted 
     */
    private void createDocuments(String inputted) {
        try {
            URL location = ForensicExpertWitnessReportConfigPanel.class.getProtectionDomain().getCodeSource().getLocation();
            TemplateOne_doc = new Document(location.getFile() + "ForensicExpertWitnessReport\\" + "1.docx"); 
            TemplateTwo_doc = new Document(location.getFile() + "ForensicExpertWitnessReport\\" + "2.docx"); 
            TemplateThree_doc = new Document(location.getFile() + "ForensicExpertWitnessReport\\" + "3.docx"); 
            inputted_doc = new Document(inputted);             
        } catch(Exception e){
            Logger.getLogger(ForensicExpertWitnessReportConfigPanel.class.getName()).log(Level.SEVERE, "Failed to create document objects", e);
        }  
    }
    
    /**
     * GetSelectedDocument Method
     * First Accessor Method.
     * 
     * Returns Document Objects of Forensic Expert Witness Reports
     * 
     * @return inputted_doc
     * @return TemplateOne_doc
     * @return TemplateTwo_doc
     * @return TemplateThree_doc
     * @return null
     */
    public Document getSelectedDocument() {
        if (selectedDocumentName.equals(inputted_name)) {
            return inputted_doc;
        }
        if (selectedDocumentName.equals(TemplateOne_name)) {
            return TemplateOne_doc;
        }
        if (selectedDocumentName.equals(TemplateTwo_name)) {
            return TemplateTwo_doc;
        }
        if (selectedDocumentName.equals(TemplateThree_name)) {
            return TemplateThree_doc;
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
     * GetEvidenceHeading Method
     * Third Accessor Method.
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
     * GetSelectedTagNames Method
     * Fourth Accessor Method.
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
     * Class TagNamesListModel of package ForensicExpertWitnessReport
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
     * Class TagsNamesListCellRenderer of package ForensicExpertWitnessReport
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
    private javax.swing.JButton selectAllButton;
    private javax.swing.JButton deselectAllButton;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JComboBox<String> expertWitnessReportComboBox;
    private javax.swing.JButton chooseExpertWitnessReportButton;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JLabel jLabel3;
    // End of variables declaration//GEN-END:variables    
    
}
