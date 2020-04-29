# SongBeamerImport

SongBeamer Import for PowerPoint (as an Add-In) and LibreOffice (as Extension).
This is realized with Macros, written in Basic (VBA).

## Platforms

The PowerPoint version has been tested with Office 2016 for Mac and Windows.
The LibreOffice version has been tested with LibreOffice 6.4.3.2 for Mac

## Sources for Macro and Extension Developers

### LibreOffice

 - https://wiki.documentfoundation.org/Main_Page
 - https://wiki.openoffice.org/wiki/Documentation/
 - https://wiki.documentfoundation.org/Documentation/Publications#Desktop_Reference_Cards
 - https://api.libreoffice.org/docs/idl/ref/index.html
 - https://www.uni-due.de/~abi070/ooo.html (http://www.pitonyak.org/OOME_3_0.pdf)

### PowerPoint Add-In

To make Buttons in PowerPoint ribbon bar, `customUI.xml` was edited, using the following editors:

 - https://bettersolutions.com/vba/ribbon/custom-ui-editor.htm
 - http://www.andypope.info/vba/ribboneditor_2010.htm
 
 and using the following sources:
 
 - http://www.pptfaq.com/FAQ01216-Create-an-ADD-IN-with-Ribbon-buttons-that-run-macros-when-clicked.htm
 - https://oerttel.net/data/documents/Tuhls-kleine-RibbonUI-Fibel.pdf
 - https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/f2a8e3c0-14cb-4ad3-88cd-a8b5b1b9a8a0
 - https://bert-toolkit.com/imagemso-list.html

to the following content of `customUI.xml`:

```XML
<!-- This adds a new group to the Design tab and adds a couple of buttons to the group -->
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">

<ribbon>
  <tabs>
    <tab idMso="TabHome">

      <group id="id_Group_SongBeamerHome"
      Label = "SongBeamer"
      insertAfterMso="GroupSlides">

        <button
          Id = "id_Button_SongBeamer_Import"
          Label = "Import Songs"
          Size = "large"
          imageMso = "TextFromFileInsert"
         OnAction = "SongBeamerImport"
         ScreenTip = "SongBeamer Song Dateien (*.sng) ausw_hlen und als PowerPoint-Folien importieren."
         />

      </group>
    </tab>
    <tab idMso="TabInsert">

      <group id="id_Group_SongBeamerInsert"
      Label = "SongBeamer"
      insertBeforeMso="GroupInsertTables">

        <button
          Id = "id_Button_SongBeamer_Import2"
          Label = "Import Songs"
          Size = "large"
          imageMso = "TextFromFileInsert"
         OnAction = "SongBeamerImport"
         ScreenTip = "SongBeamer Song Dateien (*.sng) ausw_hlen und als PowerPoint-Folien importieren."
         />

      </group>
    </tab>

  </tabs>
</ribbon>

</customUI>
```
