import re
import xml.etree.ElementTree as ET
import zipfile
import os

class DocxExtractor:
    def __init__(self, docx: str):
        self.docx_path = docx
        self._nsmap = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        self._doc_zip = self.__unzip(self.docx_path)
        self.filelist = self._doc_zip.namelist()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.closeDocx()

    def __unzip(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError(
                'File {} does not exist.'.format(file_path))

        # unzip the docx in memory
        return zipfile.ZipFile(file_path)

    def __xml2text(self, xml):
        """
            A string representing the textual content of this run, with content
            child elements like ``<w:tab/>`` translated to their Python
            equivalent.
            Adapted from: https://github.com/python-openxml/python-docx/
        """
        text = u''
        root = ET.fromstring(xml)

        for child in root.iter():
            if child.tag == self.__qn('w:t'):
                text += child.text if child.text is not None else ''
            elif child.tag == self.__qn('w:tab'):
                text += '\t'
            elif child.tag in (self.__qn('w:br'), self.__qn('w:cr')):
                text += '\n'
            elif child.tag == self.__qn("w:p"):
                text += '\n\n'
        return text

    def __qn(self, tag):
        """
            Stands for 'qualified name', a utility function to turn a namespace
            prefixed tag name into a Clark-notation qualified tag name for lxml. For
            example, ``qn('p:cSld')`` returns ``'{http://schemas.../main}cSld'``.
            Source: https://github.com/python-openxml/python-docx/
        """
        prefix, tagroot = tag.split(':')
        uri = self._nsmap[prefix]
        return '{{{}}}{}'.format(uri, tagroot)

    def getFileTree(self):
        fileTree = dict()

        for file in self.filelist:
            c_dir = fileTree
            for _dir in file.split("/"):
                c_dir[_dir] = c_dir.get(_dir, {})
                c_dir = c_dir[_dir]
        return fileTree

    def getTitle(self):
        return self._doc_zip.filename
    
    def unzipDocxToFolder(self,store_dir="",exist_ok= False):

        if(store_dir):
            if os.path.exists(store_dir):
                if not exist_ok:
                    raise FileExistsError(f"{store_dir} directory already exists")
            else:
                os.makedirs(store_dir)
        else:
            store_dir = self.getTitle()

        for file in self.filelist:
            full_store_path = os.path.join(store_dir, file)
            dir_path = os.path.dirname(full_store_path)
            print(full_store_path,dir_path)
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)
            with open(full_store_path, "wb") as fh:
                 fh.write(self._doc_zip.read(file))

    def getBodyXmlTree(self):
        return str(self._doc_zip.read('word/document.xml'),encoding="utf8")

    def extractDocumentBodyText(self):
        text = u""
        text += self.__xml2text(self._doc_zip.read('word/document.xml'))
        return re.sub(r"\n{1,}", "\n", text).strip()

    def extractHeaderText(self):
        text = u""
        xmls = 'word/header[0-9]*.xml'
        for fname in self.filelist:
            if re.match(xmls, fname):
                text += self.__xml2text(self._doc_zip.read(fname))

        return re.sub(r"\n{1,}", "\n", text).strip()

    def extractFooterText(self):
        text = u""
        xmls = 'word/footer[0-9]*.xml'
        for fname in self.filelist:
            if re.match(xmls, fname):
                text += self.__xml2text(self._doc_zip.read(fname))

        return re.sub(r"\n{1,}", "\n", text).strip()

    def extractImages(self, store_dir="", exist_ok=False):
        if os.path.exists(store_dir):
            if not exist_ok:
                raise FileExistsError(f"{store_dir} directory already exists")
        else:
            os.mkdir(store_dir, mode=0o777)

        image_files = list(self.getFileTree()["word"]["media"].keys())

        for file in image_files:
            full_store_path = os.path.join(store_dir, os.path.basename(file))
            with open(full_store_path, "wb") as fh:
                fh.write(self._doc_zip.read("word/media/"+file))

    def extractUrls(self):
        pass

    def closeDocx(self):
        self._doc_zip.close()

if __name__ == '__main__':

    with DocxExtractor(r"TWO WEEKS OF INTERNSHIP I.docx") as dh:

        print(dh.extractDocumentBodyText())

        dh.unzipDocxToFolder("wow/unzip",exist_ok=True)
