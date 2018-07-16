/*
 * Class ForensicExpertWitnessReportConfigPanelOptions.java of package ForensicExpertWitnessReport
 * 
 * Using this class you are able to display a graphical user interface (GUI)
 * inside Autopsy which allows the user to select which tagged files he or
 * she would like to report, and the forensic expert witness report he or she
 * would like to report to. This further allows the selection of three included
 * forensic expert witness report templates.
 * 
 * This class was written for a final year project for
 * the degree of Computer and Digital Forensics BSc (Hons),
 * at Northumbria University in Newcastle. This project included
 * the aim of aiding in automation, ease and increase effectiveness of digital 
 * forensic practitioners whilst conducting digital forensic
 * investigations in Autopsy.
 * 
 * @author Chris Wipat
 * @version 19.04.2018
 */

package ForensicExpertWitnessReport;
 
import java.util.ArrayList;
import javax.swing.JOptionPane;
import org.sleuthkit.autopsy.corecomponents.OptionsPanel;
import org.sleuthkit.autopsy.ingest.IngestModuleGlobalSettingsPanel;

public final class ForensicExpertWitnessReportConfigPanelOptions extends IngestModuleGlobalSettingsPanel implements OptionsPanel {
    
    // Declare instance variables.
    private String colourName;
    private String hexadecimalColourCode = "";
    private final ArrayList<String> colourNames = new ArrayList<String>();
    private final ArrayList<String> hexadecimalColourCodes = new ArrayList<String>();
    private boolean hexadecimalCode_matches_a_ComboBox_colour = false;
    private boolean inputted_code = false;
    private String last_applied = "1337";

    /**
     * Constructor for objects of class ForensicExpertWitnessReportConfigPanelOptions
     * First and only Constructor.
     * 
     * Call methods which populate GUI components, display the GUI to the user and add
     * in the colour and hexadecimal number settings which have been previously selected.
     * 
     * @param colourName
     * @param hexadecimalColourCode
     */
    public ForensicExpertWitnessReportConfigPanelOptions(String colourName, String hexadecimalColourCode) {        
        initComponents();
        populateTableColours();
        
        // If values are not null and thus the fed in paramaters are previous settings
        if (colourName != null && !hexadecimalColourCode.isEmpty())
        {
            // If hexadecimal colour code is bogus, use the distinguished colour.        
            if (hexadecimalColourCode.trim().length() != 6) {

                // Match the colour with it's corresponding hexadecimal value.
                for (int i=0; i<colourNames.size(); i++)
                {
                    if(colourName.equals(colourNames.get(i))) {
                        tableColoursComboBox.setSelectedIndex(i);
                        hexadecimalEntryField.setText(hexadecimalColourCodes.get(i));
                    }
                }

            // If hexadecimal value isn't bogus, use it & set ComboBox accordingly.
            } else {
                
                // Set hexadecimalEntryField to the given hexadecimalColourCode.
                hexadecimalEntryField.setText(hexadecimalColourCode);
                
                // Run the checker to see if the hexadecimal code matches a ComboBox colour.
                checker();
                
                // If hexadecimal code doesn't match a ComboBox colour, add "Inputted Code" to the ComboBox.
                if (hexadecimalCode_matches_a_ComboBox_colour == false) {                
                    tableColoursComboBox.addItem("Inputted Code");
                    tableColoursComboBox.setSelectedIndex(12);
                    inputted_code = true;
                }               
                  
            }
            hexadecimalCode_matches_a_ComboBox_colour = false;  
            saveSettings();
        }

    }
    
    /**
     * PopulateTableColours method
     * First Mutator Method.
     * 
     * Populate the ComboBox and ArrayLists with the names and hexadecimal codes of the table colours.
     * 
     */
    private void populateTableColours() {      
        tableColoursComboBox.removeAllItems();        
        tableColoursComboBox.addItem("Black"); colourNames.add("Black"); hexadecimalColourCodes.add("000000");
        tableColoursComboBox.addItem("Red"); colourNames.add("Red"); hexadecimalColourCodes.add("990000");
        tableColoursComboBox.addItem("Orange"); colourNames.add("Orange"); hexadecimalColourCodes.add("ff9900");
        tableColoursComboBox.addItem("Navy Blue"); colourNames.add("Navy Blue"); hexadecimalColourCodes.add("003366");
        tableColoursComboBox.addItem("Blue"); colourNames.add("Blue"); hexadecimalColourCodes.add("0039e6");
        tableColoursComboBox.addItem("Aqua"); colourNames.add("Aqua"); hexadecimalColourCodes.add("00ffff");
        tableColoursComboBox.addItem("Yellow"); colourNames.add("Yellow"); hexadecimalColourCodes.add("ffff00");
        tableColoursComboBox.addItem("Dark Green"); colourNames.add("Dark Green"); hexadecimalColourCodes.add("009933");
        tableColoursComboBox.addItem("Green"); colourNames.add("Green"); hexadecimalColourCodes.add("33cc33");
        tableColoursComboBox.addItem("Light Green"); colourNames.add("Light Green"); hexadecimalColourCodes.add("66ff66");
        tableColoursComboBox.addItem("Pink"); colourNames.add("Pink"); hexadecimalColourCodes.add("ff66ff");
        tableColoursComboBox.addItem("Purple"); colourNames.add("Purple"); hexadecimalColourCodes.add("6600ff");
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
     * On colour selected, add the hexadecimal code to JTextField1..
     * 
     * @param evt 
     */     
    private void tableColoursComboBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_tableColoursComboBoxActionPerformed
        for (int i=0; i<colourNames.size(); i++)
        {
            if(tableColoursComboBox.getSelectedItem().toString().equals(colourNames.get(i))) 
            {
                hexadecimalEntryField.setText(hexadecimalColourCodes.get(i));
            }
        }        
        
        if(tableColoursComboBox.getSelectedItem().toString().equals("Inputted Code")) {
            hexadecimalEntryField.setText(hexadecimalColourCode);
        }
    }//GEN-LAST:event_tableColoursComboBoxActionPerformed

    /**
     * Checker Method
     * Fifth Mutator Method.
     * 
     * Compares the hexadecimal code in the hexadecimalEntryField to the ComboBox colours, using our ArrayList of hexadecimal colours.       * 
     */
    private void checker() {
        for (int i=0; i<hexadecimalColourCodes.size(); i++)
        {
            if (hexadecimalColourCodes.get(i).equals(hexadecimalEntryField.getText())) {
                
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
        
       
        /// FIX THE PROBLEM !!!!!!!!!!!!!!!!!!!!!!


        // If the last applied setting is not equal to the hexadecimalEntryField (to stop applying twice)
        if (!last_applied.equals(hexadecimalEntryField.getText().trim())) {
                   
            // Run the checker to see if the hexadecimal code matches a ComboBox colour.
            checker();

            // If the hexadecimal code doesn't match a ComboBox colour, and is valid hexadecimal, then add "Inputted Code" to the ComboBox and set it to be selected.
            if (hexadecimalCode_matches_a_ComboBox_colour == false) {

                // If the hexadecimal value isn't bogus
                if (hexadecimalEntryField.getText().trim().length() == 6) {

                    // Use our test to see if "Inputted Code" has been added to the ComboBox before.
                    if (inputted_code) {
                        tableColoursComboBox.removeItemAt(12);
                    }

                    // Add Inputted Code to the ComboBox, set it to selected and set inputted_code to true.
                    tableColoursComboBox.addItem("Inputted Code");
                    tableColoursComboBox.setSelectedIndex(12);
                    inputted_code = true;

                    // Let the user know the application has been sucessful.
                    JOptionPane.showMessageDialog(null, "Sucessfully applied.", "Your settings have been applied.", JOptionPane.INFORMATION_MESSAGE);
                }

                // If the hexadecimal value is bogus, throw an error letting them know that hexadecimal codes should be 6 characters.
                else {
                     JOptionPane.showMessageDialog(null, "Hexadecimal colour codes must be 6 characters.", "Incorrect hexadecimal code.", JOptionPane.ERROR_MESSAGE);
                }            
            }

            // If the hexadecimal value does match a ComboBox colour, throw sucessfully applied.
            else {
                JOptionPane.showMessageDialog(null, "Sucessfully applied.", "Your settings have been applied.", JOptionPane.INFORMATION_MESSAGE);
            }
        
            // Set back to default
            hexadecimalCode_matches_a_ComboBox_colour = false;      
            
            // Save settings
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
        
        // If the last applied setting is not equal to the hexadecimalEntryField (to stop applying twice)
        if (!last_applied.equals(hexadecimalEntryField.getText().trim())) {
        
            // Run the checker to see if the hexadecimal code matches a ComboBox colour.
            checker();

            // If the hexadecimal code doesn't match a ComboBox colour, and is valid hexadecimal, then add "Inputted Code" to the ComboBox and set it to be selected.
            if (hexadecimalCode_matches_a_ComboBox_colour == false) {

                // If the hexadecimal value isn't bogus
                if (hexadecimalEntryField.getText().trim().length() == 6) {

                    // Use our test to see if "Inputted Code" has been added to the ComboBox before.
                    if (inputted_code) {
                        tableColoursComboBox.removeItemAt(12);
                    }

                    // Add Inputted Code to the ComboBox, set it to selected and set inputted_code to true.
                    tableColoursComboBox.addItem("Inputted Code");
                    tableColoursComboBox.setSelectedIndex(12);
                    inputted_code = true;

                }
         
            }
            
            // Set back to default
            hexadecimalCode_matches_a_ComboBox_colour = false;           

            // Save settings
            saveSettings();
            
            // Set last applied to applied settings
            last_applied = hexadecimalEntryField.getText().trim();
            
        }
    }//GEN-LAST:event_hexadecimalEntryFieldFocusLost
    
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton ApplyButton;
    private javax.swing.JLabel hexCodeLabel;
    private javax.swing.JTextField hexadecimalEntryField;
    private javax.swing.JLabel tableColourLabel;
    private javax.swing.JComboBox<String> tableColoursComboBox;
    // End of variables declaration//GEN-END:variables

    
}
