# -*- coding: utf-8 -*-
#pylint: disable=I0011,C0111,W0110,W0141
import json
import os
import tabulate
import argparse
import shutil

def _get_exercise(identifier=None, json_file=u'exercise.json'):
    json_data = None
    with open(json_file) as data_file:
        json_data = json.load(data_file)

    if identifier:
        if json_data and identifier in json_data:
            json_data = {k:v for k, v in json_data.iteritems() if k == identifier}
        else:
            json_data = {}

    return json_data or {}

def _print_exercises(exercises):
    headers = ['ID', 'Type', 'Name']

    print 'Following are the available exercise templates\n'
    exercise_list = []
    for key, values in exercises.iteritems():
        exercise_list.append([key, values['type'], values['name']])

    print tabulate.tabulate(exercise_list, headers=headers)

def _create_folder(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

def _create_tree(name):
    directory = name
    _create_folder(directory)

    directory = u'{}\\{}'.format(name, name)
    _create_folder(directory)

    directory = u'{}\\{}'.format(name, u'Solución')
    _create_folder(directory)

def __main__():
    parser = argparse.ArgumentParser(
        description='creates MAP structure from template')

    parser = argparse.ArgumentParser(description='Process some integers.')
    parser.add_argument('-l', '--list', action='store_true', dest='listOpt',
                        help='an integer for the accumulator')
    parser.add_argument('-r', '--remove', dest='removeID', metavar='ID',
                        help='removes an existing exercise structure')
    parser.add_argument('-c', '--create', dest='createID', metavar='ID',
                        help='creates a new exercise structure')
    parser.add_argument('-n', '--name', dest='exName', metavar='name',
                        help='name of the new exercise')
    parser.add_argument('-j', '--json', dest='jsonFile',
                        metavar='json', default=u'exercise.json',
                        help='json file which contains the template definitions')

    args = parser.parse_args()

    json_file = args.jsonFile or u'exercise.json'

    if args.removeID:
        exercises = _get_exercise(args.createID, json_file)
        _print_exercises(exercises)
    elif args.createID:
        exercises = _get_exercise(args.createID, json_file)

        exercise = exercises[args.createID]
        name = args.exName or exercise['name']
        typ = exercise['type'].capitalize()
        ext = u'docx' if typ == u'Word' else u'xlsx'

        _create_tree(name)

        file_path = u'{}\\Enunciado ({}).docx'.format(name, typ)
        shutil.copyfile(exercise['statement'], file_path)

        file_path = u'{}\\{}\\Doc{} (Solución).{}'.format(name, u'Solución', typ, ext)
        shutil.copyfile(exercise['solution'], file_path)

        file_path = u'{}\\{}\\Doc{}.{}'.format(name, name, typ, ext)
        shutil.copyfile(exercise['empty'], file_path)
    else:
        exercises = _get_exercise(json_file=json_file)
        _print_exercises(exercises)

__main__()
