package ForensicExpertWitnessReport;

/*
 * Class ForensicExpertWitnessReportConfigPanel.java of package ForensicExpertwitnessReport
 * 
 * Using this class you are able to ...
 * 
 * @author Chris Wipat
 * @version 02.03.2018
 */

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


public class ForensicExpertWitnessReportConfigPanel extends javax.swing.JPanel {

    // Declare instance variables
    private final Map<String, Boolean> tagNameSelections = new LinkedHashMap<String, Boolean>();
    private final TagNamesListModel tagsNamesListModel = new TagNamesListModel();  
    private final TagsNamesListCellRenderer tagsNamesRenderer = new TagsNamesListCellRenderer();
    private List<TagName> tagNames;
    private static final long serialVersionUID = 1L;   
    private String selectedDocumentName;    
    private final String TemplateOne_name = "Pre-existing Template 1";
    private final String TemplateTwo_name = "Pre-existing Template 2";
    private final String TemplateThree_name = "Pre-existing Template 3";   
    private Document TemplateOne_doc = null;
    private Document TemplateTwo_doc = null;
    private Document TemplateThree_doc = null;
    private String inputted_name;
    private String inputted_full_path;
    private Document inputted_doc = null;    
    
    ForensicExpertWitnessReportConfigPanel() {
        initComponents();
        populateTagNameComponents();
        populateForensicExpertWitnessReports();
        createDocuments("");
    }
        
    private void populateTagNameComponents() {
        
        // Get the tag names in use for the current case.
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
    
    // Populate combobox with forensic expert witness reports
    private void populateForensicExpertWitnessReports() {       
        expertWitnessReportComboBox.removeAllItems();        
        expertWitnessReportComboBox.addItem(TemplateOne_name); 
        expertWitnessReportComboBox.addItem(TemplateTwo_name);
        expertWitnessReportComboBox.addItem(TemplateThree_name);
    } 
    
    private void initComponents() {

    jScrollPane1 = new javax.swing.JScrollPane();
    tagNamesListBox = new javax.swing.JList<String>();
    selectAllButton = new javax.swing.JButton();
    deselectAllButton = new javax.swing.JButton();
    jLabel1 = new javax.swing.JLabel();
    expertWitnessReportComboBox = new javax.swing.JComboBox<String>();    
    chooseExpertWitnessReportButton = new javax.swing.JButton();
    jLabel2 = new javax.swing.JLabel(); 
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
 
    org.openide.awt.Mnemonics.setLocalizedText(jLabel1, "Export files tagged as:"); // NOI18N 
    
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
 
    org.openide.awt.Mnemonics.setLocalizedText(jLabel2, "Select Forensic Expert Witness Report"); // NOI18N
 
    javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
    this.setLayout(layout);
    layout.setHorizontalGroup(
        layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
        .addGroup(layout.createSequentialGroup()
            .addContainerGap()
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addComponent(jLabel2)
                .addComponent(jLabel1)
                .addGroup(layout.createSequentialGroup()
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jScrollPane1)
                        .addGroup(layout.createSequentialGroup()
                            .addComponent(expertWitnessReportComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, 159, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                            .addComponent(chooseExpertWitnessReportButton)))
                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                        .addComponent(deselectAllButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(selectAllButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
            .addContainerGap())
    );
    layout.setVerticalGroup(
        layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
        .addGroup(layout.createSequentialGroup()
            .addContainerGap()
            .addComponent(jLabel1)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(layout.createSequentialGroup()
                    .addComponent(selectAllButton)
                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                    .addComponent(deselectAllButton))
                .addComponent(jScrollPane1))
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
            .addComponent(jLabel2)
            .addGap(4, 4, 4)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                .addComponent(expertWitnessReportComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addComponent(chooseExpertWitnessReportButton))
            .addContainerGap())
    );
}// </editor-fold>//GEN-END:initComponents
    
    private void selectAllButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_selectAllButtonActionPerformed
        for (TagName tagName : tagNames) {
            tagNameSelections.put(tagName.getDisplayName(), Boolean.TRUE);
        }
        tagNamesListBox.repaint();
    }//GEN-LAST:event_selectAllButtonActionPerformed
 
    private void expertWitnessReportComboBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_hashSetsComboBoxActionPerformed
        selectedDocumentName = (String)expertWitnessReportComboBox.getSelectedItem();
    }//GEN-LAST:event_hashSetsComboBoxActionPerformed
 
    private void deselectAllButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_deselectAllButtonActionPerformed
        for (TagName tagName : tagNames) {
            tagNameSelections.put(tagName.getDisplayName(), Boolean.FALSE);
        }
        tagNamesListBox.repaint();
    }//GEN-LAST:event_deselectAllButtonActionPerformed

    private void chooseExpertWitnessReportButtonActionPerformed(java.awt.event.ActionEvent evt) {
    
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);  
        int result = fileChooser.showOpenDialog(null);
 
        if (result == JFileChooser.APPROVE_OPTION) {            
            inputted_full_path = fileChooser.getSelectedFile().getAbsolutePath();
            inputted_name = fileChooser.getSelectedFile().getName();
            populateForensicExpertWitnessReports();
            createDocuments(inputted_full_path);
            expertWitnessReportComboBox.addItem(inputted_name); 
        }    
    }
    
    /**
     * 
     * Mutator x
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
        return null;
   }

//    public static String getDataDir(Class c) {
//        File dir = new File(System.getProperty("user.dir"));
//        dir = new File(dir, "src");
//        dir = new File(dir, "main");
//        dir = new File(dir, "resources");
//
//        for (String s : c.getName().split("\\.")) {
//            dir = new File(dir, s);
//            if (dir.isDirectory() == false)
//                dir.mkdir();
//        }
//        System.out.println("Using data directory: " + dir.toString());
//        return dir.toString() + File.separator;
//    }
 
//  private void configureHashDatabasesButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_configureHashDatabasesButtonActionPerformed
//      HashLookupSettingsPanel configPanel = new HashLookupSettingsPanel();
//      configPanel.load();
//      if (JOptionPane.showConfirmDialog(null, configPanel, "Hash Set Configuration", JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE) == JOptionPane.OK_OPTION) {
//          configPanel.store();
//          populateHashSetComponents();
//      } else {
//          configPanel.cancel();
//          populateHashSetComponents();
//      }
//  }//GEN-LAST:event_configureHashDatabasesButtonActionPerformed
    
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

    // This class renders the items in the tag names JList component as JCheckbox components.
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
    
//    // This class renders the items in the forensic expert witness report combo list component
//    private class ComboBoxRenderer implements ListCellRenderer {
//        
//        public ComboBoxRenderer() {        
//        }
//
//        public Component getListCellRendererComponent(JList list,Object value,int index,boolean isSelected,boolean cellHasFocus) {
//            setText("asdf");
//            return this;
//        }
//        
//    }
    
//    private class Document2 extends Document {
//        private Document doc;
//        private String name = "blah";
//        
//        private Document2(String name) throws Exception {
//            this.name = name;
//            try {             
//                doc = new Document(name);
//            } catch(Exception e){
//                Logger.getLogger(ForensicExpertWitnessReport.class.getName()).log(Level.SEVERE, "Failed to create new document", e);
//            }  
//        }
//        
//        @Override
//        public String toString(){
//            return name;
//        }
//        
//        @Override
//        public Document getDocument() {
//           return doc;
//        }
//    }
    
    

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton chooseExpertWitnessReportButton;
    private javax.swing.JButton deselectAllButton;
    private javax.swing.JComboBox<String> expertWitnessReportComboBox;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JButton selectAllButton;
    private javax.swing.JList<String> tagNamesListBox = new JList<String>();
    // End of variables declaration//GEN-END:variables    
    
}
