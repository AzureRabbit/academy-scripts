# -*- coding: utf-8 -*-
###############################################################################
#    License, author and contributors information in:                         #
#    __openerp__.py file at the root folder of this module.                   #
###############################################################################

from openerp import models, fields, api
from openerp.tools.translate import _
from logging import getLogger


_logger = getLogger(__name__)


class AtTestAtQuestionRel(models.Model):
    """ The summary line for a class docstring should fit on one line.

    Fields:
      name (Char): Human readable name which will identify each record.

    """

    _name = 'at.test.at.question.rel'
    _description = (u'Relationship between at_test and at_question, this model '
                    'keeps sequence order')

    _rec_name = 'at_test_id'
    _order = 'sequence ASC'

    at_test_id = fields.Many2one(
        string='Test',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Test to which this item belongs',
        comodel_name='at.test',
        domain=[],
        context={},
        ondelete='cascade',
        auto_join=False
    )

    at_question_id = fields.Many2one(
        string='Question',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Question will be related with test',
        comodel_name='at.question',
        domain=[],
        context={},
        ondelete='cascade',
        auto_join=False
    )

    sequence = fields.Integer(
        string='Sequence',
        required=False,
        readonly=False,
        index=False,
        default=0,
        help='Question sequence order'
    )
