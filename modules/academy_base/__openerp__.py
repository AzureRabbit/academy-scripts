# -*- coding: utf-8 -*-
{
    'name': "Academy Base",

    'summary': """
        Common information and behavior used by the academy modules""",

    'description': """
        Common information and behavior used by the academy modules
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

    ],

    # always loaded
    'data': [
        'views/academy_base.xml',
    ],
    # only loaded in demonstration mode
    'demo': [
    ],
    'js': [
    ],
    'css': [
    ],
}
