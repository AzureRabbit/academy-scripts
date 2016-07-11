# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0212,E0611,C0103,R0903,C0111,F0401
###############################################################################
#    License, author and contributors information in:                         #
#    __openerp__.py file at the root folder of this module.                   #
###############################################################################

from __future__ import division

from openerp import models, fields, api
from openerp.tools.translate import _
from openerp.exceptions import ValidationError
from logging import getLogger

_logger = getLogger(__name__)


class AtQuestion(models.Model):
    """ Questions are the academy tests cornerstone. Each one of the questions
    belongs to a single topic but they can belong to more than one question in
    the selected topic.

    Fields:
      name (Char): Human readable name which will identify each record.

    """

    _name = 'at.question'
    _description = u'Referred question'

    _rec_name = 'name'
    _order = 'name ASC'

    _inherit = ['mail.thread']

    # ---------------------------- ENTITY FIELDS ------------------------------

    name = fields.Char(
        string='Name',
        required=True,
        readonly=False,
        index=True,
        default=None,
        help='Text for this question',
        size=250,
        translate=True,
        track_visibility='onchange'
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

    at_topic_id = fields.Many2one(
        string='Topic',
        required=True,
        readonly=False,
        index=False,
        default=None,
        help='Topic to which this question belongs',
        comodel_name='at.topic',
        domain=[],
        context={},
        ondelete='cascade',
        auto_join=False,
        track_visibility='onchange'
    )

    at_category_ids = fields.Many2many(
        string='Categories',
        required=True,
        readonly=False,
        index=False,
        default=None,
        help='Categories relating to this question',
        comodel_name='at.category',
        #relation='model_name_this_model_rel',
        #column1='model_name_id}',
        #column2='this_model_id',
        domain=lambda self: self._compute_at_category_ids_domain(),
        context={},
        limit=None,
        track_visibility='onchange'
    )

    at_answer_ids = fields.One2many(
        string='Answers',
        required=True,
        readonly=False,
        index=False,
        default=None,
        help='Answers will be shown as choice options for this question',
        comodel_name='at.answer',
        inverse_name='at_question_id',
        domain=[],
        context={},
        auto_join=False,
        limit=None,
        track_visibility='onchange'
    )

    at_tag_ids = fields.Many2many(
        string='Tags',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Tag can be used to better describe this question',
        comodel_name='at.tag',
        # relation='model_name_this_model_rel',
        # column1='model_name_id}',
        # column2='this_model_id',
        domain=[],
        context={},
        limit=None,
        track_visibility='onchange'
    )

    at_level_id = fields.Many2one(
        string='Difficulty level',
        required=True,
        readonly=False,
        index=False,
        default=lambda self: self._default_at_level_id(),
        help='Difficulty level of this question',
        comodel_name='at.level',
        domain=[],
        context={},
        ondelete='cascade',
        auto_join=False,
        track_visibility='onchange'
    )

    ir_attachment_ids = fields.Many2many(
        string='Attachments',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Attachments needed to solve this question',
        comodel_name='ir.attachment',
        relation='ir_attachment_this_model_rel',
        column1='ir_attachment_id',
        column2='this_model_id',
        domain=[],
        context={},
        limit=None
    )

    # ----------------------- AUXILIARY FIELD METHODS -------------------------

    def _compute_at_category_ids_domain(self):
        """ Computes domain for at_category_ids, this should allow categories
        only in the selected topic.
        """

        id_list = self.at_topic_id.at_category_ids.mapped('id')

        return [('id', 'in', tuple(id_list))]


    def _default_at_level_id(self):
        """ Computes the level_id default value
        """

        # STEP 1: Set default to none
        level_id = None

        # STEP 2: Search for all levels sorted by sequence
        at_level_domain = []
        at_level_obj = self.env['at.level']
        at_level_set = at_level_obj.search(
            at_level_domain, order="sequence ASC")

        # STEP 3: Gets the middel item from sorted set
        if at_level_set:
            middle = len(at_level_set) // 2
            level_id = at_level_set[middle].id


        return level_id

    # --------------------------- ONCHANGE EVENTS -----------------------------

    @api.onchange('at_topic_id')
    def _onchange_at_topid_id(self):
        """ Updates domain form at_category_ids, this shoud allow categories
        only in the selected topic.
        """
        domain = self._compute_at_category_ids_domain()
        _logger.debug(domain)
        return {
            'domain': {
                'at_category_ids': domain
            }
        }

    # -------------------------- PYTHON CONSTRAINS ----------------------------

    @api.one
    @api.constrains('at_answer_ids')
    def _check_at_answer_ids(self):
        """ Check if question have at last one valid answer
        """

        message = _(u'You must specify at least one correct answer')
        if not True in self.at_answer_ids.mapped('is_correct'):
            raise ValidationError(message)

    # --------------------------- SQL_CONTRAINTS ------------------------------

    _sql_constraints = [
        (
            'question_uniq',
            'UNIQUE(name)',
            _(u'There is already another question with the same name')
        )
    ]
