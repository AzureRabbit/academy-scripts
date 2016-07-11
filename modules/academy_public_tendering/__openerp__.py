# -*- coding: utf-8 -*-
{
    'name': "Academy Public Tendering",

    'summary': """
        Store and manage information about civil service entrance examination""",

    'description': """
        Store and manage information about civil service entrance examination
    """,

    'author': "Jorge Soto Garcia",
    'website': "https://github.com/sotogarcia",

    # Categories can be used to filter modules in modules listing
    # Check https://github.com/odoo/odoo/blob/master/openerp/addons/base/module/module_data.xml
    # for the full list
    'category': 'Uncategorized',
    'version': '0.1',

    # any module necessary for this one to work correctly
    'depends': [
        'academy_base'
    ],

    # always loaded
    'data': [
        'data/res_partner_category_data.xml',
        'data/res_partner_data.xml',
        'data/apt_group_data.xml',
        'data/apt_kind_data.xml',

        'views/academy_public_tendering.xml',

        'views/apt_group_view.xml',
        'views/apt_kind_view.xml',
        'views/apt_vacancy_position_view.xml',
        'views/apt_public_tendering_view.xml',
    ],
    # only loaded in demonstration mode
    'demo': [
    ],
    'js': [
    ],
    'css': [
    ],
}

