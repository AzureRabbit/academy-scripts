# -*- coding: utf-8 -*-
{
    'name': "Academy Tests",

    'summary': """
        Store and manage information about tests and their questions""",

    'description': """
        Store and manage information about tests and their questions
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
        'academy_base',
        'mail'
    ],

    # always loaded
    'data': [
        'data/at_topic_data.xml',
        'data/at_category_data.xml',
        'data/at_level_data.xml',

        'demo/at_answer_demo.xml',
        'demo/at_topic_demo.xml',
        'demo/at_category_demo.xml',
        'demo/at_question_demo.xml',
        'demo/at_tag_demo.xml',

        'security/at_answer.xml',
        'security/at_category.xml',
        'security/at_level.xml',
        'security/at_question.xml',
        'security/at_tag.xml',
        'security/at_test.xml',
        'security/at_topic.xml',

        'views/academy_tests.xml',
        'views/at_answer_view.xml',
        'views/at_category_view.xml',
        'views/at_level_view.xml',
        'views/at_question_view.xml',
        'views/at_tag_view.xml',
        'views/at_test_at_question_rel.xml',
        'views/at_test_view.xml',
        'views/at_topic_view.xml',
        'views/ir_attachment_view.xml'
    ],
    # only loaded in demonstration mode
    'demo': [
    ],
    'js': [
    ],
    'css': [
    ],
}

