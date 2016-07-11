# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0212,E0611,C0103,R0903,C0111,F0401
###############################################################################
#    License, author and contributors information in:                         #
#    __openerp__.py file at the root folder of this module.                   #
###############################################################################

from openerp import models, fields, api
from openerp.tools.translate import _
from logging import getLogger


_logger = getLogger(__name__)


class AtCategory(models.Model):
    """ Category of the question

    Fields:
      name (Char): Human readable name which will identify each record.

    """

    _name = 'at.category'
    _description = u'Category of the question'

    _rec_name = 'name'
    _order = 'sequence ASC, name ASC'

    # ---------------------------- ENTITY FIEDS -------------------------------

    name = fields.Char(
        string='Name',
        required=True,
        readonly=False,
        index=True,
        default=None,
        help='Name for this category',
        size=50,
        translate=True
    )

    description = fields.Text(
        string='Description',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Something about this category',
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

    sequence = fields.Integer(
        string='Sequence',
        required=True,
        readonly=False,
        index=False,
        default=10,
        help=('Place of this category in the order of the categories from'
              'the topic')
    )

    at_topic_id = fields.Many2one(
        string='Topic',
        required=True,
        readonly=False,
        index=False,
        default=None,
        help='Topic to which this category belongs',
        comodel_name='at.topic',
        domain=[],
        context={},
        ondelete='cascade',
        auto_join=False
    )

    at_question_ids = fields.Many2many(
        string='Questions',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Questions relating to this category',
        comodel_name='at.question',
        # relation='at_question_this_model_rel',
        # column1='at_question_id}',
        # column2='this_model_id',
        domain=[],
        context={},
        limit=None
    )

    # --------------------------- SQL_CONTRAINTS ------------------------------

    _sql_constraints = [
        (
            'categoryr_by_topic_uniq',
            'UNIQUE(at_topic_id, name)',
            _(u'There is already another category with the same name in this topic')
        )
    ]
