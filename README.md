# RichPaste
OneNote 2016 Addin - Paste Rich Text Into Your Notebook! <b><h3> To install addin - click on 'RichPasteInstall.msi' above and download</b></h3> 

<i>If you have any further questions or would like to contact me then send me an email at <a href="samueljoeoconnor@hotmail.co.uk">samueljoeoconnor@hotmail.co.uk</a> </i>

<hr>
#Tutorial

<h3><u> Contents </u></h3>

[Creating The Ribbon Button](#Creating)

[How OneNote Pages Are Structured/Stored](#structured)

[How To Use The OneNote API](#API)


<hr><a name="Creating"/>
<b><u>Creating The Ribbon Button</u></b>

To start off the addin simply follow this guide: http://www.malteahrens.com/#/blog/howto-onenote-dev/

Most of the guide is still relevant to OneNote 2016 but there are some notable changes:
<br>
<br>
1) Use .Net version 4.5 instead of the newer 4.5.2 etc (in Project Properties >> Application >> Target Framework). Your compiler will probably warn you to do this anyway.

2) When you add the assemblies you'll need the following COM assemblies (for Office 2016)

    1) Microsoft OneNote 15.0 Object Library
    2) Microsoft Office 16.0 Object Library

3) It is mentioned in the guide but you have to turn off the 'Embed Interop Types' from the OneNote Assembly (hit F4 on the reference)

4) Installer Templates aren't in the standard install of Visual Studio 2015, but you can download the templates here: https://visualstudiogallery.msdn.microsoft.com/f1cc3f3e-c300-40a7-8797-c509fb8933b9

5) [On 64bit machine] When adding the registry keys you will also have to add the keys to your WOW6432Node Folders (see RegKeysImport.reg for examples)

6) [On 64bit machine] You'll also have to change the output type on the Setup project to x64 (hit F4 on the project)

7) To build the Setup project along with the Class Libraries you'll need to go into 'Build' >> 'Configuration Manager' and tick 'Build' next to the Setup project

8) This might not be a problem for everyone but in my solution the addin would always fail on load up if I had any "Microsoft.Office.Interop.OneNote.Application" references/objects in my start up Class (the one with your ProgID and GUID). I recommend leaving them out of the startup class and just using them in classes that the startup one uses.
<br>
<hr><a name="structured"/>
<b><u>How OneNote Pages Are Structured/Stored</u></b>

OneNote pages are stored in XML. We can view and edit this XML to update our pages.

The basic structure is:


    <Page>
        <Title>
            <OE>
                <T>
                    <![CDATA[ PAGE TITLE ]]>
                </T>
            </OE>
        </Title>
        
        <Outline >
            <Position x="35.0" y="60.0"/>"                                  
            <Size width="750.75" height="13.50" />        
        
                <OEChildren>
                    <OE>
                        <T>
                            <![CDATA[ PAGE CONTENT ]]>
                        </T>
                    </OE>
                </OEChildren>
        
        </Outline>
    </Page>
    

So obviously the CDATA[] between the "Title" tags is the page's name. 

The Outline tags specify a new box of content on the page that you can move around. 

The OEChildren belong to the Outline tags that it is between, there are usually only one OEChildren tag per Outline, but OEChildren can contain as many OE tags as you'd like. 

And the OE, T and CDATA represent lines within the content box. For each line you'll have a seperate CDATA[].

Each OE contains Attributes regarding who created it, when it was created, and a unique object reference id.

This reference it contains 

    1) A GUID   (You can create GUIDs using NewGUID() method)
    2) A object number  (This increments from the number of the previous object)
    3) A notebook hex number (I think, not 100% sure but I know that this doesn't change for me on all my pages)


Special formats add in additional tags (like bullet points adds a "Bullet" tag)... but these are the basics and all you need to understand to start editing your pages.

OneNote contains your Page's text within the CDATA[] tag in HTML. This is why some Paste functions work (like copy/pasteing from MS Word) and why Rich Text doesn't (there is no standard RTF to HTML convertion done)
<br>
<hr><a name="API"/>
<b><u>How To Use The OneNote API</u></b>

Once you've got the ribbon created and understand OneNotes data structure you'll want to start querying your own OneNote. 

In your solution first create an instance of the OneNote application:

    var oneNote = new Microsoft.Office.Interop.OneNote.Application();

Notice the fully qualified name, that is because you will most likely have another type of Application in one of your assemblies and you'll have to specify that you want OneNotes API.

One tip is that you can state that you specifically want oneNotes API in a using statement:

    using Application = Microsoft.Office.Interop.OneNote.Application;

That way the first row (and any other OneNote app declarations) can be simplified down to just:

    var oneNote = new Application();
    
Now that you have an instance of the onenote API you can start using their methods.

There is a great big list of them all here: https://msdn.microsoft.com/en-us/library/office/gg649853(v=office.14).aspx

Most of them are pretty straight forward, but the few that I believe to be the most useful are actually not on the list.

I'm talking about <b>getting the current NoteBook/Section/Page</b>

These methods are a little more hidden. Here they are listed below:

    string currentOpenNoteBook = oneNote.Windows.CurrentWindow.CurrentNotebookId;
    string currentOpenSection  = oneNote.Windows.CurrentWindow.CurrentSectionId;    
    string currentOpenPage = oneNote.Windows.CurrentWindow.CurrentPageId;
    
With the above ID's you can then use the GetHierarchy()/ GetPageContent()/ UpdatePageContent() methods to edit your onenote programmatically.
