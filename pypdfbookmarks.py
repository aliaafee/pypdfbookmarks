import sys
import getopt
import csv
import code
from PyPDF2 import PdfFileReader, PdfFileWriter


class BookmarkNode(object):
    def __init__(self, title=None):
        self.parent = None
        self.children = []
        self.title = title
        self.page_number = 0
        self.tabs = None
        return


    def parse_csv(self, entry_text):
        """ Load csv line in the format
            "\t\t","Title","page_number" """
        entry_array = list(csv.reader([entry_text],
                           delimiter=",", quotechar="\"",
                           quoting=csv.QUOTE_ALL))[0]
        print(entry_array)
        self.page_number = int(entry_array[2]) - 1
        self.title = entry_array[1]
        self.tabs = re.match('\t*', entry_array[0]).group(0).count('\t')


    def add_child(self, child):
        """Add a child node of type BookmarkNode"""
        self.children.append(child)
        child.parent = self


    def set_parent(self, new_parent):
        """Set the parent of this node to a new one"""
        old_parent = self.parent
        new_parent.add_child(self)
        if old_parent is not None:
            old_parent.children.remove(self)


    def move_to(self, new_index):
        """Move this node to a new location in the list of children"""
        if self.parent is None:
            print("Cannot move the root node")
            return
        
        self.parent.children.insert(
            new_index,
            self.parent.children.pop(self.parent.children.index(self)))


    def remove(self):
        """Remove/Delete the Node"""
        if self.parent is None:
            print("Cannot remove the root node")

        self.parent.children.remove(self)


    def load_from_pdf(self, pdfreader):
        """Load bookmarks from PyPDF2 PdfFileReader"""
        def _setup_page_id_to_num(pdf, pages=None, _result=None, _num_pages=None):
            if _result is None:
                _result = {}
            if pages is None:
                _num_pages = []
                pages = pdf.trailer["/Root"].getObject()["/Pages"].getObject()
            t = pages["/Type"]
            if t == "/Pages":
                for page in pages["/Kids"]:
                    _result[page.idnum] = len(_num_pages)
                    _setup_page_id_to_num(pdf, page.getObject(), _result, _num_pages)
            elif t == "/Page":
                _num_pages.append(1)
            return _result
        
        pg_id_num_map = _setup_page_id_to_num(pdfreader)

        def _generate_tree(node, outlines):
            current_node = None
            for item in outlines:
                if type(item) is not list:
                    current_node = BookmarkNode()
                    current_node.title = item.title
                    current_node.page_number = pg_id_num_map[item.page.idnum] + 1
                    node.add_child(current_node)
                else:
                    _generate_tree(current_node, item)

        _generate_tree(self, pdfreader.outlines)


    def add_to_pdf(self, pdfwriter):
        """Save this bookmarks tree structure to PyPDF2 PdfFileWriter"""
        def _add_bookmark(node, pdfwriter, parent=None):
            pdf_node = None
            if node.parent is not None:
                pdf_node = pdfwriter.addBookmark(node.title, node.page_number - 1, parent=parent)
            for child in node.children:
                _add_bookmark(child, pdfwriter, pdf_node)
        _add_bookmark(self, pdfwriter)


    def print_tree(self, num=0, node=None, depth=0):
        """Recursively print all the nodes of this tree"""
        if node is None:
            node = self
        else:
            print("{}[{}] {}".format("   " * depth, num, node))
        for num, child in enumerate(node.children):
            self.print_tree(num, child, depth + 1)


    def print_children(self):
        """Print all the children of this node"""
        for num, child in enumerate(self.children):
            print("[{}] {}".format(num, child))


    def __repr__(self):
        """String representation of object"""
        return "{} -> p{}{}".format(
            self.title,
            self.page_number,
            ", c{}".format(len(self.children)) if self.children else "")




def write_pdf(bookmarks_tree, pdfreader, output_filename):
    """Write the pdfreader to new file using the given bookmarks tree"""
    output = PdfFileWriter()
    
    for page in pdfreader.pages:
        output.addPage(page)

    bookmarks_tree.add_to_pdf(output)

    with open(output_filename, "wb") as output_file:
        output.write(output_file)




source_filename = ""
pdfreader = None
tree = None


def load_pdf(filename):
    """Load pdf from file, PyPDF2 PdfFileReader to 'pdfreader', and
       Bookmarks to 'tree'"""
    global source_filename, pdfreader, tree
    print("Loading pdf {}...".format(filename))
    source_filename = filename
    pdfreader = PdfFileReader(filename)
    tree = BookmarkNode("Root")
    tree.load_from_pdf(pdfreader)


def save_pdf(filename):
    """Save the open pdf to new file, using the bookmarks tree in 'tree'"""
    global pdfreader, tree
    if pdfreader is None:
        print("Open a pdf document fisrt")
        return

    if not(pdfreader and tree):
        print("Load a pdf using load_pdf(filename)")
        return
    write_pdf(tree, pdfreader, filename)


def pypdfbm_help():
    print("'load_pdf(filename)' to load a pdf file")
    print("'tree' is an object of type BookmarkNode.")
    print("       dir(tree) for list of available methods.")
    print("'save_pdf(filename)' to save currently open pdf to new")
    print("       file using the bookmarks tree in 'tree'.")
    print("'exit()' to exit.")
    print("")


def shell():
    code.interact(
        banner="PyPDF Bookmarks Shell, pypdfbm_help() for help",
        local=globals()
    )


def app_license():
    """ App license """
    print("pyPDF Bookmarks")
    print("--------------")
    print("Copyright (C) 2019 Ali Aafee")
    print("")


def usage():
    """App Usage"""
    app_license()
    print("Usage: pypdfbookmarks <pdf filename>")
    print("    -h, --help")
    print("       Displays this help")
    print("")
    print("Opens an interactive shell that lets you edit bookmarks")
    print("use command 'pypdfbm_help()' for list of commands")


def main(argv):
    try:
        opts, args = getopt.getopt(argv, "h", ["help"])
    except getopt.GetoptError:
        usage()
        sys.exit(2)

    for opt, arg in opts:
        if opt in ("-h", "--help"):
            usage()
            sys.exit()

    if args:
        load_pdf(args[0])

    shell()


if __name__ == "__main__":
    main(sys.argv[1:])
