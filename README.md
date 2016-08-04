# RichPaste
OneNote 2016 Addin - Paste Rich Text Into Your Notebook!
________________________________________________________________________

# Method

<b><u>Creating The ribbon button</u></b>

To start off the addin simply follow this guide: http://www.malteahrens.com/#/blog/howto-onenote-dev/

Most of the guide is still relevant to OneNote 2016 but there are some notable changes:

1) When you add the assemblies you'll need the following COM assemblies (for Office 2016)

    1) Microsoft OneNote 15.0 Object Library
    2) Microsoft Office 16.0 Object Library

2) It is mentioned in the guide but you have to turn off the 'Embed Interop Types' from the OneNote Assembly (hit F4 on the reference)

3) Installer Templates aren't in the standard install of Visual Studio 2015, but you can download the templates here: https://visualstudiogallery.msdn.microsoft.com/f1cc3f3e-c300-40a7-8797-c509fb8933b9

4) [On 64bit machine] When adding the registry keys you will also have to add the keys to your WOW6432Node Folders (see RegKeysImport.reg for examples)

5) [On 64bit machine] You'll also have to change the output type on the Setup project to x64 (hit F4 on the project)

6) To build the Setup project along with the Class Libraries you'll need to go into 'Build' >> 'Configuration Manager' and tick 'Build' next to the Setup project

<hr>
<b><u>How OneNote pages are structured and how we can edit them</u></b>

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

The OEChildren belong to the Outline that it is between, there are usually only one OEChildren tag per Outline, but OEChildren can contain as many OE tags as you'd like. 

And the OE, T and CDATA represent lines within the content box. For each line you'll have a seperate CDATA[].

    
