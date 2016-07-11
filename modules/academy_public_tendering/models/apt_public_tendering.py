# -*- coding: utf-8 -*-
###############################################################################
#    License, author and contributors information in:                         #
#    __openerp__.py file at the root folder of this module.                   #
###############################################################################

from openerp import models, fields, api
from openerp.tools.translate import _
from logging import getLogger
from datetime import datetime, timedelta

_logger = getLogger(__name__)


class AptPublicTendering(models.Model):
    """ Public tendering information

    Fields:
      name (Char): Human readable name which will identify each record.

    """

    _name = 'apt.public.tendering'
    _description = u'Public tendering'

    _rec_name = 'name'
    _order = 'approval DESC'

    name = fields.Char(
        string='Name',
        required=True,
        readonly=False,
        index=True,
        default=None,
        help='Name for this public tendering',
        size=50,
        translate=True
    )

    description = fields.Text(
        string='Description',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Something about this public tendering',
        translate=True
    )

    active = fields.Boolean(
        string='Active',
        required=False,
        readonly=False,
        index=False,
        default=True,
        help=('If the active field is set to false, it will allow you '
              'to hide record without removing it.')
    )

    administration_id = fields.Many2one(
        string='Administration',
        required=True,
        readonly=False,
        index=False,
        default=None,
        help='Choose the administration related with tendering',
        comodel_name='res.partner',
        domain=lambda self: self.administration_id_domain(),
        context={},
        ondelete='cascade',
        auto_join=False
    )

    approval = fields.Date(
        string='Approval date',
        required=False,
        readonly=False,
        index=False,
        default=fields.Date.today(),
        help='Choose the approval date'
    )

    announcement = fields.Date(
        string='Announcement date',
        required=False,
        readonly=False,
        index=False,
        default=fields.Date.today(),
        help='Choose the Announcement date'
    )

    submissions_deadline = fields.Date(
        string='Deadline for submissions',
        required=False,
        readonly=False,
        index=False,
        default=lambda self: self.default_submissions_deadline(),
        help='Choose the last day for submission of requests'
    )

    apt_vacancy_position_ids = fields.One2many(
        string='Vacancy positions',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Add offered vacancy positions',
        comodel_name='apt.vacancy.position',
        inverse_name='apt_public_tendering_id',
        domain=[],
        context={},
        auto_join=False,
        limit=None
    )

    ir_atachment_ids = fields.Many2many(
        string='Attachments',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Documents related with the public tendering',
        comodel_name='ir.attachment',
        # relation='model_name_this_model_rel',
        # column1='model_name_id',
        # column2='this_model_id',
        domain=[],
        context={},
        limit=None
    )

    bulletin_board_url = fields.Char(
        string='Bulletin Board',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='URl of the bulletin board url',
        size=256,
        translate=True
    )

    official_journal_url = fields.Char(
        string='Official journal',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='URl of the article in the official journal',
        size=256,
        translate=True
    )

    # ----------------------- AUXILIAR FIELD METHODS --------------------------

    @staticmethod
    def default_submissions_deadline():
        """ Returns the default value for deadline_for_submissions field. This
        must be twenty days after the announcement date.
        """
        return fields.Date.to_string(
            datetime.now() + timedelta(days=20)
        )

    @api.model
    def administration_id_domain(self):
        """ Computes domain for administration_id
        """

        xid = 'academy_public_tendering.res_partner_category_civil_service'
        record = self.env.ref(xid)

        return [
            ('is_company', '=', True),
            ('category_id', '=', record.id)
        ]




