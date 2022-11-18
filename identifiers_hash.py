import enum
import glob
import hashlib
import json
import os
import shutil
import subprocess
import traceback
import olefile
import pcodedmp_extractor


VBA_IMPHASH_IDENTIFIERS = None
def load_vba_imphash_identifiers():
    global VBA_IMPHASH_IDENTIFIERS
    python_script_path = os.path.split(__file__)[0]
    predefined_identifiers_path = os.path.join(python_script_path, 'import_identifiers.json')
    with open(predefined_identifiers_path, 'r') as fh:
        VBA_IMPHASH_IDENTIFIERS = set(json.load(fh))


# called when module is imported
load_vba_imphash_identifiers()


class OfficeFileType(enum.Enum):
    OLE = 1
    OOXML = 2
    INVALID = 3


"""
Returns a (vba_imphash, list_imphash_identifiers, list_non_imphash_identifiers) tuple.
"""
def compute_imphash(file_path):
    office_file_type = _get_office_file_type(file_path)
    
    if office_file_type == OfficeFileType.OLE:
        return _compute_imphash_for_ole_office_file(file_path)
    elif office_file_type == OfficeFileType.OOXML:
        return _compute_imphash_for_ooxml_office_file(file_path)

    return 'INVALID_OFFICE_FILE', [], []


def _get_office_file_type(office_file_path):
    with open(office_file_path, 'rb') as fh:
        file_header = fh.read(2)
    
    if file_header == b'\xd0\xcf':
        return OfficeFileType.OLE
    elif file_header == b'PK':
        return OfficeFileType.OOXML
    
    return OfficeFileType.INVALID


"""
Returns a (vba_imphash, list_imphash_identifiers, list_non_imphash_identifiers) tuple.
"""
def _compute_imphash_for_ole_office_file(office_file_path):
    print(f'Computing import hash for OLE office file {office_file_path}.')

    if not _is_ole_office_file_valid(office_file_path):
        return 'INVALID_OLE_OFFICE_FILE', [], []
    
    vba_project_stream = _read_vba_project_stream_for_ole_office_file(office_file_path)
    return compute_imphash_from_vba_project_stream(vba_project_stream)


def _is_ole_office_file_valid(office_file_path):
    if not _is_ole_file(office_file_path):
        print(f'File {office_file_path} is not a valid OLE file.')
        return False

    if not _ole_office_file_has_vba_macros_storage(office_file_path):
        print(f'File {office_file_path} does not have the "Macros\\VBA" storage.')
        return False
    
    if not _ole_office_file_has_vba_project_stream(office_file_path):
        print(f'File {office_file_path} does not have the "Macros\\VBA\\_VBA_PROJECT" stream.')
        return False
    
    return True


def _is_ole_file(file_path):
    try:
        return olefile.isOleFile(file_path)
    except:
        print(f'[Exception] Path = {file_path}. Trace = {traceback.format_exc()}')
        return False


def _ole_office_file_has_vba_macros_storage(office_file_path):
    try:
        with olefile.OleFileIO(office_file_path) as ole:
            return ole.exists('macros/vba')
    except:
        print(f'[Exception] Path = {office_file_path}. Trace = {traceback.format_exc()}')
        return False


def _ole_office_file_has_vba_project_stream(office_file_path):
    try:
        with olefile.OleFileIO(office_file_path) as ole:
            return ole.exists('macros/vba/_vba_project')
    except:
        print(f'[Exception] Path = {office_file_path}. Trace = {traceback.format_exc()}')
        return False


"""
Returns a bytes object.
"""
def _read_vba_project_stream_for_ole_office_file(office_file_path):
    try:
        with olefile.OleFileIO(office_file_path) as ole:
            return ole.openstream('macros/vba/_vba_project').read()
    except:
        print(f'[Exception] Path = {office_file_path}. Trace = {traceback.format_exc()}')
        return b''

 
"""
'vba_project_stream' parameter is a bytes object.
Returns a (vba_imphash, list_imphash_identifiers, list_non_imphash_identifiers) tuple.
"""
def compute_imphash_from_vba_project_stream(vba_project_stream):
    all_identifiers = pcodedmp_extractor.get_all_identifiers(vba_project_stream)
    list_imphash_identifiers, list_non_imphash_identifiers = \
        _get_lists_categorized_identifiers(all_identifiers)
    vba_imp_hash = _compute_vba_imphash_from_identifiers(list_imphash_identifiers)
    return vba_imp_hash, list_imphash_identifiers, list_non_imphash_identifiers


"""
Returns a (list_imphash_identifiers, list_non_imphash_identifiers) tuple.
"""
def _get_lists_categorized_identifiers(all_identifiers):
    list_imphash_identifiers = []
    list_non_imphash_identifiers = []

    for identifier in all_identifiers:
        if _is_import_related_identifier(identifier):
            list_imphash_identifiers.append(identifier)
        else:
            list_non_imphash_identifiers.append(identifier)
    
    return list_imphash_identifiers, list_non_imphash_identifiers


"""
'Import' identifiers where determined from this documentation - 
https://learn.microsoft.com/en-us/office/vba/api/overview/language-reference
Haven't tested all of them. I just copied them from the docs. Some might not necessarily
appear as identifiers in Office files or might be rarely found.
"""
def _is_import_related_identifier(identifier):
    return identifier.lower() in VBA_IMPHASH_IDENTIFIERS


def _compute_vba_imphash_from_identifiers(list_imphash_identifiers):
    if len(list_imphash_identifiers) == 0:
        return 'NO_IMPHASH_IDENTIFIERS'
    import_identifiers_strs = '-'.join(list_imphash_identifiers)
    return hashlib.md5(import_identifiers_strs.encode('utf-8')).hexdigest()


"""
Returns a (vba_imphash, list_imphash_identifiers, list_non_imphash_identifiers) tuple.
"""
def _compute_imphash_for_ooxml_office_file(office_file_path):
    print(f'Computing import hash for OOXML office file {office_file_path}.')

    if not _is_ooxml_office_file_valid(office_file_path):
        return 'INVALID_OOXML_OFFICE_FILE', [], []
    
    vba_project_stream = _read_vba_project_stream_for_ooxml_office_file(office_file_path)
    return compute_imphash_from_vba_project_stream(vba_project_stream)


def _is_ooxml_office_file_valid(office_file_path):
    if not _is_ooxml_office_file(office_file_path):
        print(f'File {office_file_path} is not a valid OOXML office file.')
        return False

    if not _ooxml_file_has_vbaproject(office_file_path):
        print(f'File {office_file_path} does not have the "vbaProject.bin" OLE file embedded '
            'inside the zip.')
        return False
    
    if not _ooxml_file_has_vba_project_stream(office_file_path):
        print(f'File {office_file_path} does not have the "vbaProject.bin\\VBA\\_VBA_PROJECT" '
            'stream.')
        return False
    
    return True


def _is_ooxml_office_file(file_path):
    with open(file_path, 'rb') as fh:
        if fh.read(2) != b'PK':
            return False

    stdout_for_7z_list_command = _call_7z_list_contents_pk_zip_file(file_path)
    return '[Content_Types].xml' in stdout_for_7z_list_command


def _call_7z_list_contents_pk_zip_file(pk_file_path):
    completed_proc_info = subprocess.run(['7z', 'l', '-y', pk_file_path], capture_output=True)
    return str(completed_proc_info.stdout)


"""
The name of the OLE file found in the PK archive is usually 'vbaProject.bin',
Example files where the stream has a different name - 'EKGbZQJlSu.bin'
https://www.malware-traffic-analysis.net/2022/06/08/index.html
(D74C9EBF3A09DF2FCCD47265DDAB693862B09A4D1CFEA336675BAFF32BC83C93)
"""
def _ooxml_file_has_vbaproject(ooxml_file_path):
    stdout_for_7z_list_command = _call_7z_list_contents_pk_zip_file(ooxml_file_path)
    return '.bin' in stdout_for_7z_list_command


def _ooxml_file_has_vba_project_stream(ooxml_file_path):
    vbaProjectBin_path = os.path.join(os.path.split(ooxml_file_path)[0], 'vbaProject.bin')
    _remove_existing_vbaProjectBin_file(vbaProjectBin_path)

    _extract_vbaprojectbin_from_ooxml_file(ooxml_file_path)
    try:
        return _vbaProjectBin_file_has_vba_project_stream(vbaProjectBin_path)
    except:
        print(f'[Exception] Path = {ooxml_file_path}. Trace = {traceback.format_exc()}')
        return False


def _remove_existing_vbaProjectBin_file(vbaProjectBin_path):
    try:
        os.remove(vbaProjectBin_path)
    except:
        pass


def _extract_vbaprojectbin_from_ooxml_file(ooxml_file_path):
    workingdir_path = os.path.split(ooxml_file_path)[0]

    working_dir_temp_path = os.path.join(workingdir_path, '__TEMP__')
    _create_temp_dir_for_extracting_vbaprojectbin(working_dir_temp_path)

    _call_7z_to_extract_bin_files(ooxml_file_path, working_dir_temp_path)
    _rename_extracted_vbaprojectbin_file(working_dir_temp_path)
    shutil.rmtree(working_dir_temp_path)


def _create_temp_dir_for_extracting_vbaprojectbin(working_dir_temp_path):
    try:
        shutil.rmtree(working_dir_temp_path)
        # Removing directory if it already exists
    except FileNotFoundError:
        pass
    os.mkdir(working_dir_temp_path)


def _call_7z_to_extract_bin_files(ooxml_file_path, working_dir_temp_path):
    subprocess.run(['7z', 'e', ooxml_file_path, '-y', '*.bin', '-r', 
        f'-o{working_dir_temp_path}'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def _rename_extracted_vbaprojectbin_file(working_dir_temp_path):
    list_bin_file_names = glob.glob('*.bin', root_dir=working_dir_temp_path)

    if len(list_bin_file_names) == 0:
        return
    
    vbaProjectBin_file_name = _get_extracted_vbaprojectbin_file_name(list_bin_file_names)
    vbaProjectBin_file_path = os.path.join(working_dir_temp_path, vbaProjectBin_file_name)
    destination_path_vbaProjectBin = os.path.join(os.path.split(working_dir_temp_path)[0], 
        'vbaProject.bin')
    shutil.copy(vbaProjectBin_file_path, destination_path_vbaProjectBin)


def _get_extracted_vbaprojectbin_file_name(list_bin_file_names):
    vbaProjectBin_file_name = list_bin_file_names[0]
    list_bin_file_names_lower = [x.lower() for x in list_bin_file_names]
    if (len(list_bin_file_names) > 1) and ('vbproject.bin' in list_bin_file_names_lower):
        idx_vbaProjectBin = list_bin_file_names_lower.index('vbproject.bin')
        vbaProjectBin_file_name = list_bin_file_names[idx_vbaProjectBin]
    return vbaProjectBin_file_name


def _vbaProjectBin_file_has_vba_project_stream(vbaProjectBin_path):
    if not os.path.isfile(vbaProjectBin_path):
        return False

    with olefile.OleFileIO(vbaProjectBin_path) as ole:
        return ole.exists('vba/_vba_project')


"""
Returns a bytes object.
"""
def _read_vba_project_stream_for_ooxml_office_file(office_file_path):
    vbaProjectBin_path = os.path.join(os.path.split(office_file_path)[0], 'vbaProject.bin')

    _remove_existing_vbaProjectBin_file(vbaProjectBin_path)

    _extract_vbaprojectbin_from_ooxml_file(office_file_path)

    try:
        vba_project_buffer = None
        with olefile.OleFileIO(vbaProjectBin_path) as ole:
            vba_project_buffer = ole.openstream('vba/_vba_project').read()
        _remove_existing_vbaProjectBin_file(vbaProjectBin_path)
        return vba_project_buffer
    except:
        _remove_existing_vbaProjectBin_file(vbaProjectBin_path)
        print(f'[Exception] Path = {office_file_path}. Trace = {traceback.format_exc()}')
        return b''
