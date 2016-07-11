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


class AtAnswer(models.Model):
    """ Answer for a question

    Fields:
      name (Char): Human readable name which will identify each record.

    """

    _name = 'at.answer'
    _description = u'Answer for question'

    _rec_name = 'name'
    _order = 'sequence ASC'

    _inherit = ['mail.thread']

    # ---------------------------- ENTITY FIELDS ------------------------------

    name = fields.Char(
        string='Name',
        required=True,
        readonly=False,
        index=True,
        default=None,
        help='Text for this answer',
        size=250,
        translate=True
    )

    description = fields.Text(
        string='Description',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Something about this topic',
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

    at_question_id = fields.Many2one(
        string='Question',
        required=True,
        readonly=False,
        index=False,
        default=None,
        help='Question to which this answer belongs',
        comodel_name='at.question',
        domain=[],
        context={},
        ondelete='cascade',
        auto_join=False
    )

    is_correct = fields.Boolean(
        string='Is correct?',
        required=False,
        readonly=False,
        index=False,
        default=False,
        help='Checked means this is a right answer for the question'
    )

    sequence = fields.Integer(
        string='Sequence',
        required=True,
        readonly=False,
        index=False,
        default=10,
        help='Preference order for this answer'
    )

    # --------------------------- SQL_CONTRAINTS ------------------------------

    _sql_constraints = [
        (
            'answer_by_question_uniq',
            'UNIQUE(name, at_question_id)',
            _(u'There is already another answer with the same text')
        )
    ]

    # --------------------------- PUBLIC METHODS ------------------------------

    @api.multi
    def cmd_open_in_form(self):
        return {
            'name': 'Answers',
            'view_type': 'form',
            "view_mode": 'form',
            'res_model': 'at.answer',
            'type': 'ir.actions.act_window',
            'res_id': self.id,
            'target': 'current',
            'state': 'paid'
        }


