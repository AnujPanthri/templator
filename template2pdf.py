import zipfile
from copy import deepcopy
# import xml.etree.ElementTree as ET
# from xml.dom import minidom

from lxml import etree as ET


import tempfile
import os,shutil
from docx2pdf import convert

# print(ET.fromstring(open('p.xml').read()))
# print(open('Web/word/document.xml').read())
# print(ET.fromstring(open('Web/word/document.xml').read().encode()))
# exit()

# Microsoft's XML makes heavy use of XML namespaces; thus, we'll need to reference that in our code
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}



class docx_writer:
    def __init__(self,filename):
        self.zip1=zipfile.ZipFile(filename)
        self.tmp_dir = tempfile.mkdtemp()
        # print(self.tmp_dir)
        self.zip1.extractall(self.tmp_dir)
        self.root=ET.fromstring(open(os.path.join(self.tmp_dir,'word/document.xml'), 'r').read().encode())
        self.replica_elements=[]
        self._join_tags()
    
    

    def _itertext(self):
        """Iterator to go through xml tree's text nodes"""
        for node in self.root.iter(tag=ET.Element):
            if self._check_element_is(node, 't'):
                yield (node, node.text)

    def _check_element_is(self, element, type_char):
        word_schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        return element.tag == '{%s}%s' % (word_schema,type_char)


    def _join_tags(self):
        chars = []
        openbrac = False
        inside_openbrac_node = False

        for node,text in self._itertext():
            # Scan through every node with text
            for i,c in enumerate(text):
                # Go through each node's text character by character
                if c == '[':
                    openbrac = True # Within a tag
                    inside_openbrac_node = True # Tag was opened in this node
                    openbrac_node = node # Save ptr to open bracket containing node
                    chars = []
                elif c== ']':
                    assert openbrac
                    if inside_openbrac_node:
                        # Open and close inside same node, no need to do anything
                        pass
                    else:
                        # Open bracket in earlier node, now it's closed
                        # So append all the chars we've encountered since the openbrac_node '['
                        # to the openbrac_node
                        chars.append(']')
                        openbrac_node.text += ''.join(chars)
                        # Also, don't forget to remove the characters seen so far from current node
                        node.text = text[i+1:]
                        # stored reference of replica element
                        self.replica_elements.append((openbrac_node,openbrac_node.text))
                    openbrac = False
                    inside_openbrac_node = False
                else:
                    # Normal text character
                    if openbrac and inside_openbrac_node:
                        # No need to copy text
                        pass
                    elif openbrac and not inside_openbrac_node:
                        chars.append(c)
                    else:
                        # outside of a open/close
                        pass
            if openbrac and not inside_openbrac_node:
                # Went through all text that is part of an open bracket/close bracket
                # in other nodes
                # need to remove this text completely
                node.text = ""
            inside_openbrac_node = False
        
    def save_xml(self,filename='person.xml'):
        # tree = ET.ElementTree(self.root)
        
        with open(filename, 'w') as f:
            f.write(ET.tostring(self.root, pretty_print=True, xml_declaration = True, encoding='UTF-8', standalone=True).decode())



    def save_docx(self,output_filename):
            """ Create a temp directory, expand the original docx zip.
                Write the modified xml to word/document.xml
                Zip it up as the new docx
            """

            with open(os.path.join(self.tmp_dir,'word/document.xml'), 'w') as f:
                xmlstr = ET.tostring(self.root, pretty_print=True, xml_declaration = True, encoding='UTF-8', standalone=True).decode()
                # print(xmlstr)
                f.write(xmlstr)


            # # Get a list of all the files in the original docx zipfile
            filenames = self.zip1.namelist()
            # print(filenames)
            # # Now, create the new zip file and add all the filex into the archive
            zip_copy_filename = output_filename
            with zipfile.ZipFile(zip_copy_filename, "w") as docx:
                for fname in filenames:
                    docx.write(os.path.join(self.tmp_dir,fname), fname)
        
    def close(self):
        # Clean up the temp dir
        shutil.rmtree(self.tmp_dir)

    def update_xml(self,detail_dict):
        for node,txt in self.replica_elements:
            
            if(txt[1:-1] in detail_dict.keys()):
                print("'{}' replaced".format(txt[1:-1]))
                node.text=detail_dict[txt[1:-1]]   


filename='Web Front.docx'
doc1=docx_writer(filename)

detail_dict={"name":"vasu",'subject':"Science","prof":"Mr.GIG Mann","subjectcode":"420"}

for name in ["Anuj","Vasu","Bagga","Messi"]:
    detail_dict['name']=name
    output_filename="{}_test.docx".format(detail_dict["name"])
    doc1.update_xml(detail_dict)
    doc1.save_docx(output_filename)
    convert(output_filename)
    os.remove(output_filename)

doc1.close()