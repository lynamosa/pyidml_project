import zipfile

from pyidml.opc.serialized import PackageReader

# if __name__ == '__main__':
# fpath = 'data/Coffee-Obsession_content.idml'
fpath = 'data/_default.idml'
idml: PackageReader = PackageReader(fpath)
idml.root.stories
# idml.graphic.add_rgb([34, 34, 34])
# idml.graphic.add_rgb([122, 122, 122])
# idml.root.add_spread
idml.save('data/addcolor.idml')
with zipfile.ZipFile('data/addcolor.idml', "r") as zip_ref:
    zip_ref.extractall('data/addcolor/')
