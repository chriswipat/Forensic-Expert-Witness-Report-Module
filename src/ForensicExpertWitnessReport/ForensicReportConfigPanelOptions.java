/*
 * Class ForensicReportConfigPanelOptions.java of package ForensicReport
 * 
 * Using this class you are able to display a graphical user interface (GUI)
 * inside Autopsy which allows the user to select which tagged files he or
 * she would like to report, and the forensic expert witness report he or she
 * would like to report to. This further allows the selection of three included
 * forensic expert witness report templates.
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
 
import java.util.ArrayList;
import javax.swing.JOptionPane;
import org.sleuthkit.autopsy.corecomponents.OptionsPanel;
import org.sleuthkit.autopsy.ingest.IngestModuleGlobalSettingsPanel;

public final class ForensicReportConfigPanelOptions extends IngestModuleGlobalSettingsPanel implements OptionsPanel {
    
    // Declare instance variables.
    private String colourName;
    private String hexadecimalColourCode = "";
    private final ArrayList<String> colourNames = new ArrayList<String>();
    private final ArrayList<String> hexadecimalColourCodes = new ArrayList<String>();
    private boolean hexadecimalCode_matches_a_ComboBox_colour = false;
    private String inputted_code = "";
    private String last_applied = "1337";
    private int ComboBoxUsage = 0;

    /**
     * Constructor for objects of class ForensicReportConfigPanelOptions
     * First and only Constructor.
     * 
     * Call methods which populate GUI components, display the GUI to the user and add
     * in the colour and hexadecimal number settings which have been previously selected.
     * 
     * @param colourName
     * @param hexadecimalColourCode
     */
    public ForensicReportConfigPanelOptions(String colourName, String hexadecimalColourCode) {        
        initComponents();
        populateComboBoxColours();
        populateArrayListColours();
        
        // If values are not empty and thus the fed in paramaters are previous settings, do the following
        if (!colourName.isEmpty() || !hexadecimalColourCode.isEmpty()) {
            
            // If hexadecimal value isn't bogus, (can be a hexadecimal value fed in with a genuine colour) use it & set ComboBox accordingly.
            if (hexadecimalColourCode.trim().length() == 6) {
                
                // Set hexadecimalEntryField to the given hexadecimalColourCode.
                hexadecimalEntryField.setText(hexadecimalColourCode);
                
                // Run the checker to see if the hexadecimal code matches a ComboBox colour.
                checker();
                
                // If hexadecimal code doesn't match a ComboBox colour, add "Inputted Code" to the ComboBox.
                if (hexadecimalCode_matches_a_ComboBox_colour == false) {                
                    inputted_code = hexadecimalColourCode;
                    tableColoursComboBox.addItem("Inputted Code");
                    tableColoursComboBox.setSelectedIndex(12);                    
                }  

            // If hexadecimal colour code is bogus, match & use the distinguished colour name.    
            } else {
                
                // Match the colour with it's corresponding hexadecimal value & use it.
                for (int i=0; i<colourNames.size(); i++)
                {
                    if(colourNames.get(i).equals(colourName)) {
                        tableColoursComboBox.setSelectedIndex(i);
                        hexadecimalEntryField.setText(hexadecimalColourCodes.get(i));
                    }
                }

            }
            
            hexadecimalCode_matches_a_ComboBox_colour = false;  
            saveSettings();
        }

    }
    
    /**
     * PopulateComboBox method
     * First Mutator Method.
     * 
     * Populate the ComboBox with the names and hexadecimal codes of the table colours.
     * 
     */
    private void populateComboBoxColours() {      
        tableColoursComboBox.removeAllItems();        
        tableColoursComboBox.addItem("Black"); 
        tableColoursComboBox.addItem("Red"); 
        tableColoursComboBox.addItem("Orange"); 
        tableColoursComboBox.addItem("Navy Blue"); 
        tableColoursComboBox.addItem("Blue"); 
        tableColoursComboBox.addItem("Aqua"); 
        tableColoursComboBox.addItem("Yellow"); 
        tableColoursComboBox.addItem("Dark Green"); 
        tableColoursComboBox.addItem("Green"); 
        tableColoursComboBox.addItem("Light Green"); 
        tableColoursComboBox.addItem("Pink"); 
        tableColoursComboBox.addItem("Purple"); 
    }
    
    /**
     * PopulateArrayLists method
     * Second Mutator Method.
     * 
     * Populate the ArrayLists with the names and hexadecimal codes of the table colours.
     * 
     */
    private void populateArrayListColours() { 
        
        colourNames.clear();
        hexadecimalColourCodes.clear();  
        
        colourNames.add("Black"); hexadecimalColourCodes.add("000000");
        colourNames.add("Red"); hexadecimalColourCodes.add("990000");
        colourNames.add("Orange"); hexadecimalColourCodes.add("e68a00");
        colourNames.add("Navy Blue"); hexadecimalColourCodes.add("003366");
        colourNames.add("Blue"); hexadecimalColourCodes.add("000099");
        colourNames.add("Aqua"); hexadecimalColourCodes.add("00ffff");
        colourNames.add("Yellow"); hexadecimalColourCodes.add("ffff00");
        colourNames.add("Dark Green"); hexadecimalColourCodes.add("009933");
        colourNames.add("Green"); hexadecimalColourCodes.add("33cc33");
        colourNames.add("Light Green"); hexadecimalColourCodes.add("66ff66");
        colourNames.add("Pink"); hexadecimalColourCodes.add("ff66ff");
        colourNames.add("Purple"); hexadecimalColourCodes.add("6600ff");
        
        if (!inputted_code.isEmpty()) 
        {
            colourNames.add("Inputted Code");
        }
    }
    
    /**
     * Load method
     * Second Mutator Method.
     * 
     * Load the configuration panel. Required while implementing OptionsPanel.
     * 
     */
    @Override
    public void load() {
        
    }

    /**
     * Store method
     * Third Mutator Method.
     * 
     * Store the user settings. Required while implementing OptionsPanel.
     * 
     */
    @Override
    public void store() {
        saveSettings();
    }
    
    /**
     * Store method
     * Fourth Mutator Method.
     * 
     * Save the user settings. Required while implementing OptionsPanel.
     * 
     */
    @Override
    public void saveSettings() {
        colourName = tableColoursComboBox.getSelectedItem().toString();
        hexadecimalColourCode = hexadecimalEntryField.getText();
    }
    
    /**
     * GetColourName Method
     * First Accessor Method.
     * 
     * Return the name of the colour.
     * 
     * @return colourName
     */
    public String getTableColourName() {
        return colourName;
    }    
    
     /**
     * GetHexadecimalColourCode Method
     * Second Accessor Method.
     * 
     * Return the hexadecimal value of the colour.
     * 
     * @return hexadecimalColourCode
     */
    public String getHexadecimalColourCode() {
        return hexadecimalColourCode;
    }
    
     /**
     * GetColourNames Method
     * Third Accessor Method.
     * 
     * Return the ArrayList of colour names.
     * 
     * @return colourNames
     */
    public ArrayList<String> getColourNames() {
        return colourNames;
    }
    
     /**
     * GetInputtedTableColour Method
     * Fourth Accessor Method.
     * 
     * Return the ArraList of colour hexadecimal codes.
     * 
     * @return hexadecimalColourCodes
     */
    public ArrayList<String> getColourHexadecimals() {
        return hexadecimalColourCodes;
    }
    
    /**
     * InitComponents method
     * Third Mutator Method
     * 
     * Set the GUI of every component and display the GUI to the user.
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
     */
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        tableColourLabel = new javax.swing.JLabel();
        tableColoursComboBox = new javax.swing.JComboBox<String>();
        hexCodeLabel = new javax.swing.JLabel();
        ApplyButton = new javax.swing.JButton();
        hexadecimalEntryField = new javax.swing.JTextField();

        tableColourLabel.setText("Table Colour:");

        tableColoursComboBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                tableColoursComboBoxActionPerformed(evt);
            }
        });

        hexCodeLabel.setText("Hex Code:");

        ApplyButton.setText("Apply");
        ApplyButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ApplyButtonActionPerformed(evt);
            }
        });

        hexadecimalEntryField.setText("000000");
        hexadecimalEntryField.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusLost(java.awt.event.FocusEvent evt) {
                hexadecimalEntryFieldFocusLost(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(tableColoursComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(tableColourLabel))
                .addGap(30, 30, 30)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(hexCodeLabel)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(hexadecimalEntryField, javax.swing.GroupLayout.PREFERRED_SIZE, 78, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addComponent(ApplyButton)))
                .addContainerGap(16, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(tableColourLabel)
                    .addComponent(hexCodeLabel))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(tableColoursComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(ApplyButton)
                    .addComponent(hexadecimalEntryField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(17, Short.MAX_VALUE))
        );
    }// </editor-fold>//GEN-END:initComponents


     /**
     * TableColoursComboBoxActionPerformed Method
     * Fourth Mutator Method.
     * 
     * On colour selected, add the corresponding hexadecimal code to JTextField1.
     * 
     * @param evt 
     */     
    private void tableColoursComboBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_tableColoursComboBoxActionPerformed
        
        populateArrayListColours();
        ComboBoxUsage++;
        
        if (tableColoursComboBox.getSelectedIndex() < 12) {
            hexadecimalEntryField.setText(hexadecimalColourCodes.get(tableColoursComboBox.getSelectedIndex()));
        } else {
            hexadecimalEntryField.setText(inputted_code);
        }
 
    }//GEN-LAST:event_tableColoursComboBoxActionPerformed

    /**
     * Checker Method
     * Fifth Mutator Method.
     * 
     * Compares the hexadecimal code in the hexadecimalEntryField to the ComboBox colours, using our ArrayList of hexadecimal colours.       
     */
    private void checker() {
        for (int i=0; i<hexadecimalColourCodes.size(); i++)
        {
            if (hexadecimalColourCodes.get(i).equals(hexadecimalEntryField.getText().trim())) {
                
                //If colour found, set a variable and set the corresponding ComboBox index.
                hexadecimalCode_matches_a_ComboBox_colour = true;
                tableColoursComboBox.setSelectedIndex(i);
                
            }
        }
    }
    
     /**
     * ApplyButtonActionPerformed Method
     * Sixth Mutator Method.
     * 
     * On apply pressed, save the settings if the hexadecimal value is an inputted one.
     * 
     * @param evt 
     */ 
    private void ApplyButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ApplyButtonActionPerformed
        
        // If the last applied setting is not equal to the hexadecimalEntryField (to stop applying twice)
        if (!last_applied.equals(hexadecimalEntryField.getText().trim())) {
                              
                // If the hexadecimal value isn't bogus
                if (hexadecimalEntryField.getText().trim().length() == 6) {
                    
                // Run the checker to see if the hexadecimal code matches a ComboBox colour.
                checker();

                // If the hexadecimal code doesn't match a ComboBox colour, and is valid hexadecimal, then add "Inputted Code" to the ComboBox and set it to be selected.
                if (hexadecimalCode_matches_a_ComboBox_colour == false) {

                    // Use our test to see if "Inputted Code" has been added to the ComboBox before.
                    if (!inputted_code.isEmpty()) {
                        tableColoursComboBox.removeItemAt(12);
                    }
                    
                    // Set inputted_code to the valid hexadecimal code.
                    inputted_code = hexadecimalEntryField.getText();
                    
                    // Add Inputted Code to the ComboBox & set it to selected
                    tableColoursComboBox.addItem("Inputted Code");
                    tableColoursComboBox.setSelectedIndex(12);

                    // Let the user know the application has been successful.
                    JOptionPane.showMessageDialog(null, "Your settings have been saved.", "Colour applied", JOptionPane.INFORMATION_MESSAGE);
                }
                
                // If the hexadecimal value does match a ComboBox colour, and CombBox isn't as default, throw sucessfully saved. (Avoid saving settings on fresh load)
                else {
                    if(ComboBoxUsage > 2 ){
                        JOptionPane.showMessageDialog(null, "Your settings have been saved.", "Colour applied", JOptionPane.INFORMATION_MESSAGE);
                    }
                }
          
            }
                
            // If the hexadecimal value is bogus, throw an error letting them know that hexadecimal codes should be 6 characters.
            else {
                 JOptionPane.showMessageDialog(null, "Hexadecimal colour codes must be 6 characters.", "Incorrect hexadecimal code.", JOptionPane.ERROR_MESSAGE);
            }      
        
            // Set back to default
            hexadecimalCode_matches_a_ComboBox_colour = false;
            
            // Save settings regardless
            saveSettings();
            
             // Set last applied to applied settings
            last_applied = hexadecimalEntryField.getText().trim();
        }
        
    }//GEN-LAST:event_ApplyButtonActionPerformed

     /**
     * HexadecimalEntryFieldFocusLost Method
     * Seventh Mutator Method.
     * 
     * On hexadecimalEntryField input, check for user inputted hexadecimal and respond accordingly.
     * 
     * @param evt 
     */ 
    private void hexadecimalEntryFieldFocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_hexadecimalEntryFieldFocusLost
    }//GEN-LAST:event_hexadecimalEntryFieldFocusLost
    
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton ApplyButton;
    private javax.swing.JLabel hexCodeLabel;
    private javax.swing.JTextField hexadecimalEntryField;
    private javax.swing.JLabel tableColourLabel;
    private javax.swing.JComboBox<String> tableColoursComboBox;
    // End of variables declaration//GEN-END:variables

    
}
