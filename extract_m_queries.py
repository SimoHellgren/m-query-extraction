from zipfile import ZipFile
import xml.etree.ElementTree as ET
import base64
import struct
from io import BytesIO

def extract_m_queries(stream: BytesIO) -> bytes:
    '''Extracts M queries (Power Query) from a given xlsx file.
       First opens the file as a zip archive and looks for a specific file
       containing another compressed archive, containing (among other things)
       the file in which the queries are stored.
    '''
    archive = ZipFile(stream)

    # find the file containing m queries (the 'item' file containing a DataMashup-element)
    items = filter(
        lambda x: x.filename.startswith('customXml/item'),
        archive.filelist
    )

    roots = (ET.parse(archive.open(file)).getroot() for file in items)
    tree = next(root for root in roots if 'DataMashup' in root.tag)

    # decode contents and extract the OPC package binary. For details, see:
    # https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/27b1dd1e-7de8-45d9-9c84-dfcc7a802e37
    binary = base64.b64decode(tree.text)
    _, package_parts_length = struct.unpack('II', binary[:8])
    package_parts_binary = binary[8:8+package_parts_length]

    # construct zipfile and return the contents of the file we're interested in
    package = ZipFile(BytesIO(package_parts_binary))

    return package.read('Formulas/Section1.m')


if __name__ == "__main__":
    import sys

    for file in sys.argv[1:]:
        with open(file, "rb") as f:
            data = extract_m_queries(f)
            print(data)
