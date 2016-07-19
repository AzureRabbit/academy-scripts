# -*- coding: utf-8 -*-
###############################################################################
#    License, author and contributors information in:                         #
#    __openerp__.py file at the root folder of this module.                   #
###############################################################################
#pylint: disable=I0011,W0141

import random
from datetime import timedelta

class RandomValues(object):
    """ Generates random information will be used in tests
    """

    _person_list = (
        u'Pilar Lopez Perez',
        u'Jose Luis Mendez Diaz',
        u'Jaime Abilleira Moldes',
        u'Jorge Soto Garcia'
    )

    _subject_list = (
        u'Oposiciones',
        u'Ofimatica',
        u'Informatica',
        u'Burotica',
        u'Texts',
        u'Simulacros',
        u'Ejercicios'
    )

    _topic_list = (
        u'Fuente',
        u'Parrafo',
        u'Borde',
        u'Tabulacion',
        u'Columna',
        u'Capital',
        u'Nota',
        u'Numeracion',
        u'Viñeta',
        u'Tabla',
        u'Imagen',
        u'Estilo',
        u'Encabezado',
        u'Pie',
        u'Forma',
        u'Digujo',
        u'Marco',
        u'Campo',
    )

    _company_list = [
        u'Academia',
        u'Academia Postal',
        u'Academia Postal 3',
        u'Postal 3',
        u'Postal'
    ]

    _unspecified_list = [
        u'No ha{0} sido indicad{1}{2}',
        u'No ha{0} sido establecid{1}{2}',
        u'No ha{0} sido especificad{1}{2}',
        u'Aparece{0} vaci{1}{2}',
    ]

    instance = None

    def __new__(cls, *args, **kargs):
        if cls.instance is None:
            cls.instance = object.__new__(cls, *args, **kargs)
        return cls.instance

    def __init__(self):
        pass

    @staticmethod
    def _ensure_iterable(value):
        """ Ensures value is iterable
        """
        result = []

        if value:
            result = value if hasattr(value, '__iter__') else [value]

        return result

    def values(self, values, exclude=None, number=1, name='values'):
        """ Return one or several random values
        @param number (int): number of values will be returned. Default 1
        """
        # STEP 1: Get a list with all allowed values
        exclude = self._ensure_iterable(exclude)
        exclude = [k.upper() for k in exclude]
        allowed = [k for k in values if k.upper() not in exclude]

        # STEP 2: Check if there are enough values to return
        assert number >= 0 and number <= len(allowed), \
            u'There are not enough %s to return' % name

        # STEP 3: Return a list with required random values
        return random.sample(allowed, number)

    def person(self, exclude=None, number=1):
        """ Return one or several random persons
        @param number (int): number of persons will be returned. Default 1
        """
        return self.values(self._person_list, exclude, number, u'persons')

    def subject(self, exclude=None, number=1):
        """ Return one or several random subjects
        @param number (int): number of subjects will be returned. Default 1
        """
        return self.values(self._subject_list, exclude, number, u'subjects')

    def topic(self, exclude=None, number=1):
        """ Return one or several random topics
        @param number (int): number of topics will be returned. Default 1
        """
        return self.values(self._topic_list, exclude, number, u'topics')

    def company(self, exclude=None, number=1):
        """ Return one or several random topics
        @param number (int): number of topics will be returned. Default 1
        """
        return self.values(self._company_list, exclude, number, u'topics')

    def integer(self, minn=0, maxn=10, exclude=None, number=1):
        """ Return one or several random integers

        @minn: minimun valid integer
        @maxn: maximun valid integer
        @exclude: values to exclude
        @param number (int): number of topics will be returned. Default 1
        """

        # STEP 1: Get a list with all allowed values
        exclude = self._ensure_iterable(exclude)
        numbers = range(minn, maxn)
        allowed = [k for k in numbers if k not in exclude]

        # STEP 2: Check if there are enough values to return
        assert number >= 0 and number <= len(allowed), \
            u'There are not enough %s to return' % 'integers'

        # STEP 3: Return a list with required random values
        return random.sample(allowed, number)

    def date(self, date_base, offset=1, exclude=None, number=1):
        """ Return one or several random integers

        @date_base: date will be used as base, offset will be computed in
        positive and nevative from this date.
        @offset: maximun offset. It will be used as positive an nevative values.
        @exclude: dates to exclude
        @param number (int): number of topics will be returned. Default 1
        """

        # STEP 1: Offset with negative values and positive values
        numbers = range(-offset, +offset)
        exclude = self._ensure_iterable(exclude)

        # STEP 2: Get a list with all allowed values
        allowed = []
        for number in numbers:
            new_date = date_base + timedelta(days=number)
            if new_date not in exclude:
                allowed.append(new_date)

        # STEP 3: Check if there are enough values to return
        assert number >= 0 and number <= len(allowed), \
            u'There are not enough %s to return' % 'dates'

        # STEP 4: Return random values from allowed
        return random.sample(allowed, number)

    def unspecified(self, female=False, plural=False):
        """ Return random text will be user for not specified properties
        @param propperty: property name
        """

        pattern = self.values(
            self._unspecified_list, [], 1, u'unspecified texts')

        genre = u'a' if female else u'o'
        plurals = [u'n', u's'] if plural else [u'', u'']

        result = pattern[0].format(plurals[0], genre, plurals[1])

        return result




