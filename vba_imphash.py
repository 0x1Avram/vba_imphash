import collections
import json
import os
import shutil
import sys
import identifiers_hash


def show_cmdline_usage():
    script_name = os.path.split(__file__)[1]
    print(f'Usage: \n'
        f'1) For extracting the vba_imphash and displaying the identifiers for a single file: '
        f'{script_name} file_path\n'
        f'****Example****: {script_name} "details 07.20.doc.old"\n\n\n'
        f'2) For clustering files based on the computed vba_imphash:\n'
        f'  a) Without creating the clusters on disk: '
        f'{script_name} unclustered_files_path\n'
        f'  ****Example****: {script_name} "/home/test/Unclustered files/"\n\n'
        f'  b) Creating the clusters on disk: '
        f'{script_name} unclustered_files_path clusters_destination_path\n'
        f'  ****Example****: {script_name} "/home/test/Unclustered files/" '
        f'"/home/test/Clusters/"\n\n\n'
        f'7z needs to be installed and available as a command.\n'
        f'In case the clustering files version of the command line is used, the script creates '
        f'the following .json files containing relevant information in the current working '
        f'directory:  "vba_imphash_clusters.json", "imphash_identifiers.json", '
        f'"non_imphash_identifiers.json".')


def extract_vba_imphash_from_single_file(file_path):
    vba_imphash, list_imphash_identifiers, list_non_imphash_identifiers = \
        identifiers_hash.compute_imphash(file_path)
    print(f'Import identifiers: {list_imphash_identifiers}.\n'
        f'NON-Import identifiers: {list_non_imphash_identifiers}.\n'
        f'VBA import hash = {vba_imphash}.')


def cluster_office_files_directory(dir_path):
    dict_clusters = {}
    dict_imphash_identifiers = {}
    dict_non_imphash_identifiers = {}

    for file_path in _get_file_names_in_path(dir_path):
        print('')
        vba_imphash, list_imphash_identifiers, list_non_imphash_identifiers = \
            identifiers_hash.compute_imphash(file_path)
        _display_info_about_vba_imphash(dict_clusters, vba_imphash, list_imphash_identifiers, 
            file_path)
        _update_dict_clusters(dict_clusters, vba_imphash, file_path)
        _update_dict_identifiers(dict_imphash_identifiers, list_imphash_identifiers)
        _update_dict_identifiers(dict_non_imphash_identifiers, list_non_imphash_identifiers)
    
    _display_dict_clusters(dict_clusters)
    _save_dicts_to_disk(dict_clusters, dict_imphash_identifiers, dict_non_imphash_identifiers)


def _get_file_names_in_path(dir_path):
    for file_name in os.listdir(dir_path):
        file_path = os.path.join(dir_path, file_name)
        if not os.path.isfile(file_path):
            continue
        yield file_path


def _display_info_about_vba_imphash(dict_clusters, vba_imphash, list_imphash_identifiers, 
        file_path):
    if vba_imphash in dict_clusters:
        return
    print(f'File {file_path} has the vba imphash {vba_imphash} from the identifiers '
        f'{list_imphash_identifiers}.')


def _update_dict_clusters(dict_clusters, vba_imphash, file_path):
    if vba_imphash not in dict_clusters:
        dict_clusters[vba_imphash] = [file_path]
    else:
        dict_clusters[vba_imphash].append(file_path)


def _update_dict_identifiers(dict_identifiers, list_identifiers_found):
    for identifier in list_identifiers_found:
        if identifier not in dict_identifiers:
            dict_identifiers[identifier] = 1
        else:
            dict_identifiers[identifier] += 1


def _display_dict_clusters(dict_clusters):
    print('\n' * 3 + '*' * 100)
    dict_clusters = dict(sorted(dict_clusters.items(), key=lambda item: len(item[1])))
    i = 0
    for cluster_name, list_files_in_cluster_paths in dict_clusters.items():
        i += 1
        list_file_names = [os.path.split(x)[1] for x in list_files_in_cluster_paths]
        print(f'{i}) Cluster {cluster_name}. Len = {len(list_file_names)}.\n'
            f'Files: {list_file_names}')


def _save_dicts_to_disk(dict_clusters, dict_imphash_identifiers, dict_non_imphash_identifiers):
    _save_dict_clusters_to_disk(dict_clusters)
    _save_dict_identifiers_to_disk(dict_imphash_identifiers, 'imphash_identifiers.json')
    _save_dict_identifiers_to_disk(dict_non_imphash_identifiers, 'non_imphash_identifiers.json')


def _save_dict_clusters_to_disk(dict_clusters):
    dict_clusters = dict(sorted(dict_clusters.items(), key=lambda item: len(item[1])))
    dict_clusters_ordered = collections.OrderedDict()
    for cluster_name, list_files in dict_clusters.items():
        dict_clusters_ordered[cluster_name] = list_files
    _save_object_to_json_file(dict_clusters, 'vba_imphash_clusters.json')


def _save_object_to_json_file(object_to_save, json_file_path):
    with open(json_file_path, 'w') as fh:
        json.dump(object_to_save, fh, indent=4)


def _save_dict_identifiers_to_disk(dict_identifiers, dict_file_name):
    dict_identifiers = dict(sorted(dict_identifiers.items(), key=lambda item: item[1]))
    dict_identifiers_ordered = collections.OrderedDict()
    for identifier, nr in dict_identifiers.items():
        dict_identifiers_ordered[identifier] = nr
    _save_object_to_json_file(dict_identifiers_ordered, dict_file_name)


def create_clusters_on_disk(clusters_dest_path):
    dict_clusters = _load_json_from_disk('vba_imphash_clusters.json')
    for cluster_name, list_files in dict_clusters.items():
        _create_single_cluster(clusters_dest_path, cluster_name, list_files)


def _load_json_from_disk(json_file_path):
    with open(json_file_path, 'r') as fh:
        return json.load(fh)


def _create_single_cluster(clusters_dest_path, cluster_name, list_files_paths_in_cluster):
    nr_files_in_cluster = len(list_files_paths_in_cluster)
    cluster_name = f'{str(nr_files_in_cluster).zfill(5)}_{cluster_name}'
    cluster_path = os.path.join(clusters_dest_path, cluster_name)
    os.mkdir(cluster_path)
    for file_path in list_files_paths_in_cluster:
        shutil.copy(file_path, cluster_path)


def main():
    nr_args = len(sys.argv)
    if (nr_args == 1) or (nr_args > 3):
        show_cmdline_usage()
        return
    
    if nr_args == 2:
        if os.path.isfile(sys.argv[1]):
            extract_vba_imphash_from_single_file(file_path=sys.argv[1])
        else:
            cluster_office_files_directory(dir_path=sys.argv[1])
    elif nr_args == 3:
        cluster_office_files_directory(dir_path=sys.argv[1])
        create_clusters_on_disk(clusters_dest_path=sys.argv[2])
    

if __name__ == '__main__':
    main()
