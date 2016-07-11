# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0212,E0611,C0103,R0903,C0111,F0401
###############################################################################
#    License, author and contributors information in:                         #
#    __openerp__.py file at the root folder of this module.                   #
###############################################################################

from openerp import models, fields
from openerp.tools.translate import _
from logging import getLogger


_logger = getLogger(__name__)


class AtTag(models.Model):
    """ Tag can be used to better describe this question

    Fields:
      name (Char): Human readable name which will identify each record.

    """

    _name = 'at.tag'
    _description = u'Tag can be used to better describe question'

    _rec_name = 'name'
    _order = 'name ASC'

    name = fields.Char(
        string='Name',
        required=True,
        readonly=False,
        index=True,
        default=None,
        help='Name for this tag',
        size=50,
        translate=True
    )

    description = fields.Text(
        string='Description',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Something about this question',
        translate=True
    )

    active = fields.Boolean(
        string='Active',
        required=False,
        readonly=False,
        index=False,
        default=True,
        help=('If the active field is set to false, it will allow you to '
              'hide record without removing it.')
    )

    at_question_ids = fields.Many2many(
        string='Questions',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Questions relating to this tag',
        comodel_name='at.question',
        #relation='at_question_this_model_rel',
        #column1='at_question_id}',
        #column2='this_model_id',
        domain=[],
        context={},
        limit=None
    )

    # --------------------------- SQL_CONTRAINTS ------------------------------

    _sql_constraints = [
        (
            'tag_uniq',
            'UNIQUE(name)',
            _(u'There is already another tag with the same name')
        )
    ]
