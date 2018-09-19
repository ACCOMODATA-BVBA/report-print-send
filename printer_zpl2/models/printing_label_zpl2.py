# -*- coding: utf-8 -*-
# Copyright (C) 2016 SYLEAM (<http://www.syleam.fr>)
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import datetime
import logging
import time

from odoo import api, exceptions, fields, models
from odoo.tools.safe_eval import safe_eval
from odoo.tools.translate import _

_logger = logging.getLogger(__name__)

try:
    import zpl2
except ImportError:
    _logger.debug('Cannot `import zpl2`.')


class PrintingLabelZpl2(models.Model):
    _name = 'printing.label.zpl2'
    _description = 'ZPL II Label'

    name = fields.Char(required=True, help='Label Name.')
    description = fields.Char(help='Long description for this label.')
    model_id = fields.Many2one(
        comodel_name='ir.model', string='Model', required=True,
        help='Model used to print this label.')
    origin_x = fields.Integer(
        required=True, default=10,
        help='Origin point of the contents in the label, X coordinate.')
    origin_y = fields.Integer(
        required=True, default=10,
        help='Origin point of the contents in the label, Y coordinate.')
    width = fields.Integer(
        required=True, default=480,
        help='Width of the label, will be set on the printer before printing.')
    component_ids = fields.One2many(
        comodel_name='printing.label.zpl2.component', inverse_name='label_id',
        string='Label Components',
        help='Components which will be printed on the label.')
    restore_saved_config = fields.Boolean(
        string="Restore printer's configuration",
        help="Restore printer's saved configuration and end of each label ",
        default=True)
    multilabel = fields.Boolean(
        string='Multilabel',
        help='Print multiple labels at once.')
    multilabel_count = fields.Integer(
        default=1,
        help='Number of labels to print at once')
    multilabel_offset_x = fields.Integer(
        help='X coordinate offset between each occurence.')
    multilabel_offset_y = fields.Integer(
        help='Y coordinate offset between each occurence.')

    multilabel = fields.Boolean(
        string='Multilabel',
        help='Print multiple labels at once.')
    multilabel_count = fields.Integer(
        default=1,
        help='Number of labels to print at once')
    multilabel_offset_x = fields.Integer(
        help='X coordinate offset between each occurence.')
    multilabel_offset_y = fields.Integer(
        help='Y coordinate offset between each occurence.')

    @api.multi
    def _generate_zpl2_components_data(
            self, label_data, record, page_number=1, page_count=1,
            label_offset_x=0, label_offset_y=0, **extra):
        self.ensure_one()

        # Add all elements to print in a list of tuples :
        #   [(component, data, offset_x, offset_y)]
        to_print = []
        for component in self.component_ids:
            eval_args = extra
            eval_args.update({
                'object': record,
                'page_number': str(page_number + 1),
                'page_count': str(page_count),
                'time': time,
                'datetime': datetime,
            })
            data = safe_eval(component.data, eval_args) or ''

            # Generate a list of elements if the component is repeatable
            for idx in range(
                    component.repeat_offset,
                    component.repeat_offset + component.repeat_count):
                printed_data = data
                # Pick the right value if data is a collection
                if isinstance(data, (list, tuple, set, models.BaseModel)):
                    # If we reached the end of data, quit the loop
                    if idx >= len(data):
                        break

                    # Set the real data to display
                    printed_data = data[idx]

                position = idx - component.repeat_offset
                to_print.append((
                    component, printed_data,
                    label_offset_x + component.repeat_offset_x * position,
                    label_offset_y + component.repeat_offset_y * position,
                ))

        for (component, data, offset_x, offset_y) in to_print:
            component_offset_x = component.origin_x + offset_x
            component_offset_y = component.origin_y + offset_y
            if component.component_type == 'text':
                barcode_arguments = dict([
                    (field_name, component[field_name])
                    for field_name in [
                        zpl2.ARG_FONT,
                        zpl2.ARG_ORIENTATION,
                        zpl2.ARG_HEIGHT,
                        zpl2.ARG_WIDTH,
                        zpl2.ARG_REVERSE_PRINT,
                        zpl2.ARG_IN_BLOCK,
                        zpl2.ARG_BLOCK_WIDTH,
                        zpl2.ARG_BLOCK_LINES,
                        zpl2.ARG_BLOCK_SPACES,
                        zpl2.ARG_BLOCK_JUSTIFY,
                        zpl2.ARG_BLOCK_LEFT_MARGIN,
                    ]
                ])
                label_data.font_data(
                    component_offset_x, component_offset_y,
                    barcode_arguments, data)
            elif component.component_type == 'rectangle':
                label_data.graphic_box(
                    component_offset_x, component_offset_y, {
                        zpl2.ARG_WIDTH: component.width,
                        zpl2.ARG_HEIGHT: component.height,
                        zpl2.ARG_THICKNESS: component.thickness,
                        zpl2.ARG_COLOR: component.color,
                        zpl2.ARG_ROUNDING: component.rounding,
                    })
            elif component.component_type == 'circle':
                label_data.graphic_circle(
                    component_offset_x, component_offset_y, {
                        zpl2.ARG_DIAMETER: component.width,
                        zpl2.ARG_THICKNESS: component.thickness,
                        zpl2.ARG_COLOR: component.color,
                    })
            elif component.component_type == 'sublabel':
                component_offset_x += component.sublabel_id.origin_x
                component_offset_y += component.sublabel_id.origin_y
                component.sublabel_id._generate_zpl2_components_data(
                    label_data, data,
                    label_offset_x=component_offset_x,
                    label_offset_y=component_offset_y)
            else:
                barcode_arguments = dict([
                    (field_name, component[field_name])
                    for field_name in [
                        zpl2.ARG_ORIENTATION,
                        zpl2.ARG_CHECK_DIGITS,
                        zpl2.ARG_HEIGHT,
                        zpl2.ARG_INTERPRETATION_LINE,
                        zpl2.ARG_INTERPRETATION_LINE_ABOVE,
                        zpl2.ARG_SECURITY_LEVEL,
                        zpl2.ARG_COLUMNS_COUNT,
                        zpl2.ARG_ROWS_COUNT,
                        zpl2.ARG_TRUNCATE,
                        zpl2.ARG_MODULE_WIDTH,
                        zpl2.ARG_BAR_WIDTH_RATIO,
                    ]
                ])
                label_data.barcode_data(
                    component.origin_x + offset_x,
                    component.origin_y + offset_y,
                    component.component_type, barcode_arguments, data)

    @api.multi
    def _generate_zpl2_data(self, records, page_count=1, **extra):
        self.ensure_one()
        label_data = zpl2.Zpl2()

        def do_label_end():
            if self.restore_saved_config:
                label_data.configuration_update(zpl2.CONF_RECALL_LAST_SAVED)
            label_data.label_end()

        multilabel_index = 0

        for record in records:
            for page_number in range(page_count):
                if multilabel_index == 0:
                    # Initialize printer's configuration
                    label_data.label_start()
                    label_data.print_width(self.width)
                    label_data.label_encoding()

                origin_x = (
                        self.origin_x + self.multilabel_offset_x
                        * multilabel_index
                )
                origin_y = (
                        self.origin_y + self.multilabel_offset_y
                        * multilabel_index
                )
                label_data.label_home(origin_x, origin_y)

                self._generate_zpl2_components_data(
                    label_data, record, page_number=page_number,
                    page_count=page_count)

                if self.multilabel:
                    multilabel_index = \
                        (multilabel_index + 1) % self.multilabel_count

                # Restore printer's configuration and end of the label
                # before we can advance to the next length of labels
                if multilabel_index == 0:
                    do_label_end()

        # ensure the label is properly closed in case we finished printing
        # without completely filling the label length
        if multilabel_index != 0:
            do_label_end()

        return label_data.output()

    @api.multi
    def print_label(self, printer, records, page_count=1, **extra):
        for label in self:
            if records._name != label.model_id.model:
                raise exceptions.UserError(
                    _('This label cannot be used on {model}').format(
                        model=records._name))

            # Send the label to printer
            label_contents = label._generate_zpl2_data(
                records, page_count=page_count, **extra)
            printer.print_document(None, label_contents, 'raw')

        return True
