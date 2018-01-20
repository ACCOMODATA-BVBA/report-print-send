# -*- coding: utf-8 -*-
# Copyright (C) 2016 SYLEAM (<http://www.syleam.fr>)
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

from PIL import Image

# Constants for the printer configuration management
CONF_RELOAD_FACTORY = u'F'
CONF_RELOAD_NETWORK_FACTORY = u'N'
CONF_RECALL_LAST_SAVED = u'R'
CONF_SAVE_CURRENT = u'S'

# Command arguments names
ARG_FONT = u'font'
ARG_HEIGHT = u'height'
ARG_WIDTH = u'width'
ARG_ORIENTATION = u'orientation'
ARG_THICKNESS = u'thickness'
ARG_BLOCK_WIDTH = u'block_width'
ARG_BLOCK_LINES = u'block_lines'
ARG_BLOCK_SPACES = u'block_spaces'
ARG_BLOCK_JUSTIFY = u'block_justify'
ARG_BLOCK_LEFT_MARGIN = u'block_left_margin'
ARG_CHECK_DIGITS = u'check_digits'
ARG_INTERPRETATION_LINE = u'interpretation_line'
ARG_INTERPRETATION_LINE_ABOVE = u'interpretation_line_above'
ARG_STARTING_MODE = u'starting_mode'
ARG_SECURITY_LEVEL = u'security_level'
ARG_COLUMNS_COUNT = u'columns_count'
ARG_ROWS_COUNT = u'rows_count'
ARG_TRUNCATE = u'truncate'
ARG_MODE = u'mode'
ARG_MODULE_WIDTH = u'module_width'
ARG_BAR_WIDTH_RATIO = u'bar_width_ratio'
ARG_REVERSE_PRINT = u'reverse_print'
ARG_IN_BLOCK = u'in_block'
ARG_COLOR = u'color'
ARG_ROUNDING = u'rounding'
ARG_DIAMETER = u'diameter'

# Boolean values
BOOL_YES = u'Y'
BOOL_NO = u'N'

# Orientation values
ORIENTATION_NORMAL = u'N'
ORIENTATION_ROTATED = u'R'
ORIENTATION_INVERTED = u'I'
ORIENTATION_BOTTOM_UP = u'B'

# Justify values
JUSTIFY_LEFT = u'L'
JUSTIFY_CENTER = u'C'
JUSTIFY_JUSTIFIED = u'J'
JUSTIFY_RIGHT = u'R'

# Font values
FONT_DEFAULT = u'0'
FONT_9X5 = u'A'
FONT_11X7 = u'B'
FONT_18X10 = u'D'
FONT_28X15 = u'E'
FONT_26X13 = u'F'
FONT_60X40 = u'G'
FONT_21X13 = u'H'

# Color values
COLOR_BLACK = u'B'
COLOR_WHITE = u'W'

# Barcode types
BARCODE_CODE_11 = u'code_11'
BARCODE_INTERLEAVED_2_OF_5 = u'interleaved_2_of_5'
BARCODE_CODE_39 = u'code_39'
BARCODE_CODE_49 = u'code_49'
BARCODE_PDF417 = u'pdf417'
BARCODE_EAN_8 = u'ean-8'
BARCODE_UPC_E = u'upc-e'
BARCODE_CODE_128 = u'code_128'
BARCODE_EAN_13 = u'ean-13'


class Zpl2(object):
    """ ZPL II management class
    Allows to generate data for Zebra printers
    """

    def __init__(self):
        self.encoding = 'utf-8'
        self.initialize()

    def initialize(self):
        self._buffer = []

    def output(self):
        """ Return the full contents to send to the printer """
        return u'\n'.encode(self.encoding).join(self._buffer)

    def _enforce(self, value, minimum=1, maximum=32000):
        """ Returns the value, forced between minimum and maximum """
        return min(max(minimum, value), maximum)

    def _write_command(self, data):
        """ Adds a complete command to buffer """
        self._buffer.append(unicode(data).encode(self.encoding))

    def _generate_arguments(self, arguments, kwargs):
        """ Generate a zebra arguments from an argument names list and a dict of
        values for these arguments
        @param arguments : list of argument names, ORDER MATTERS
        @param kwargs : list of arguments values
        """
        command_arguments = []

        # Add all arguments in the list, if they exist
        for argument in arguments:
            if kwargs.get(argument, None) is not None:
                if isinstance(kwargs[argument], bool):
                    kwargs[argument] = kwargs[argument] and BOOL_YES or BOOL_NO
                command_arguments.append(kwargs[argument])

        # Return a zebra formatted string, with a comma between each argument
        return u','.join(map(str, command_arguments))

    def print_width(self, label_width):
        """ Defines the print width setting on the printer """
        self._write_command(u'^PW%d' % label_width)

    def configuration_update(self, active_configuration):
        """ Set the active configuration on the printer """
        self._write_command(u'^JU%s' % active_configuration)

    def label_start(self):
        """ Adds the label start command to the buffer """
        self._write_command(u'^XA')

    def label_encoding(self):
        """ Adds the label encoding command to the buffer
        Fixed value defined to UTF-8
        """
        self._write_command(u'^CI28')

    def label_end(self):
        """ Adds the label start command to the buffer """
        self._write_command(u'^XZ')

    def label_home(self, left, top):
        """ Define the label top left corner """
        self._write_command(u'^LH%d,%d' % (left, top))

    def _field_origin(self, right, down):
        """ Define the top left corner of the data, from the top left corner of
        the label
        """
        return u'^FO%d,%d' % (right, down)

    def _font_format(self, font_format):
        """ Send the commands which define the font to use for the current data
        """
        arguments = [ARG_FONT, ARG_HEIGHT, ARG_WIDTH]

        # Add orientation in the font name (only place where there is
        # no comma between values)
        font_format[ARG_FONT] += font_format.get(
            ARG_ORIENTATION, ORIENTATION_NORMAL)

        # Check that the height value fits in the allowed values
        if font_format.get(ARG_HEIGHT) is not None:
            font_format[ARG_HEIGHT] = self._enforce(
                font_format[ARG_HEIGHT], minimum=10)

        # Check that the width value fits in the allowed values
        if font_format.get(ARG_WIDTH) is not None:
            font_format[ARG_WIDTH] = self._enforce(
                font_format[ARG_WIDTH], minimum=10)

        # Generate the ZPL II command
        return u'^A' + self._generate_arguments(arguments, font_format)

    def _field_block(self, block_format):
        """ Define a maximum width to print some data """
        arguments = [
            ARG_BLOCK_WIDTH,
            ARG_BLOCK_LINES,
            ARG_BLOCK_SPACES,
            ARG_BLOCK_JUSTIFY,
            ARG_BLOCK_LEFT_MARGIN,
        ]
        return u'^FB' + self._generate_arguments(arguments, block_format)

    def _barcode_format(self, barcodeType, barcode_format):
        """ Generate the commands to print a barcode
        Each barcode type needs a specific function
        """
        def _code11(**kwargs):
            arguments = [
                ARG_ORIENTATION,
                ARG_CHECK_DIGITS,
                ARG_HEIGHT,
                ARG_INTERPRETATION_LINE,
                ARG_INTERPRETATION_LINE_ABOVE,
            ]
            return u'1' + self._generate_arguments(arguments, kwargs)

        def _interleaved2of5(**kwargs):
            arguments = [
                ARG_ORIENTATION,
                ARG_HEIGHT,
                ARG_INTERPRETATION_LINE,
                ARG_INTERPRETATION_LINE_ABOVE,
                ARG_CHECK_DIGITS,
            ]
            return u'2' + self._generate_arguments(arguments, kwargs)

        def _code39(**kwargs):
            arguments = [
                ARG_ORIENTATION,
                ARG_CHECK_DIGITS,
                ARG_HEIGHT,
                ARG_INTERPRETATION_LINE,
                ARG_INTERPRETATION_LINE_ABOVE,
            ]
            return u'3' + self._generate_arguments(arguments, kwargs)

        def _code49(**kwargs):
            arguments = [
                ARG_ORIENTATION,
                ARG_HEIGHT,
                ARG_INTERPRETATION_LINE,
                ARG_STARTING_MODE,
            ]
            # Use interpretation_line and interpretation_line_above to generate
            # a specific interpretation_line value
            if kwargs.get(ARG_INTERPRETATION_LINE) is not None:
                if kwargs[ARG_INTERPRETATION_LINE]:
                    if kwargs[ARG_INTERPRETATION_LINE_ABOVE]:
                        # Interpretation line after
                        kwargs[ARG_INTERPRETATION_LINE] = u'A'
                    else:
                        # Interpretation line before
                        kwargs[ARG_INTERPRETATION_LINE] = u'B'
                else:
                    # No interpretation line
                    kwargs[ARG_INTERPRETATION_LINE] = u'N'
            return u'4' + self._generate_arguments(arguments, kwargs)

        def _pdf417(**kwargs):
            arguments = [
                ARG_ORIENTATION,
                ARG_HEIGHT,
                ARG_SECURITY_LEVEL,
                ARG_COLUMNS_COUNT,
                ARG_ROWS_COUNT,
                ARG_TRUNCATE,
            ]
            return u'7' + self._generate_arguments(arguments, kwargs)

        def _ean8(**kwargs):
            arguments = [
                ARG_ORIENTATION,
                ARG_HEIGHT,
                ARG_INTERPRETATION_LINE,
                ARG_INTERPRETATION_LINE_ABOVE,
            ]
            return u'8' + self._generate_arguments(arguments, kwargs)

        def _upce(**kwargs):
            arguments = [
                ARG_ORIENTATION,
                ARG_HEIGHT,
                ARG_INTERPRETATION_LINE,
                ARG_INTERPRETATION_LINE_ABOVE,
                ARG_CHECK_DIGITS,
            ]
            return u'9' + self._generate_arguments(arguments, kwargs)

        def _code128(**kwargs):
            arguments = [
                ARG_ORIENTATION,
                ARG_HEIGHT,
                ARG_INTERPRETATION_LINE,
                ARG_INTERPRETATION_LINE_ABOVE,
                ARG_CHECK_DIGITS,
                ARG_MODE,
            ]
            return u'C' + self._generate_arguments(arguments, kwargs)

        def _ean13(**kwargs):
            arguments = [
                ARG_ORIENTATION,
                ARG_HEIGHT,
                ARG_INTERPRETATION_LINE,
                ARG_INTERPRETATION_LINE_ABOVE,
            ]
            return u'E' + self._generate_arguments(arguments, kwargs)

        barcodeTypes = {
            BARCODE_CODE_11: _code11,
            BARCODE_INTERLEAVED_2_OF_5: _interleaved2of5,
            BARCODE_CODE_39: _code39,
            BARCODE_CODE_49: _code49,
            BARCODE_PDF417: _pdf417,
            BARCODE_EAN_8: _ean8,
            BARCODE_UPC_E: _upce,
            BARCODE_CODE_128: _code128,
            BARCODE_EAN_13: _ean13,
        }
        return u'^B' + barcodeTypes[barcodeType](**barcode_format)

    def _barcode_field_default(self, barcode_format):
        """ Add the data start command to the buffer """
        arguments = [
            ARG_MODULE_WIDTH,
            ARG_BAR_WIDTH_RATIO,
        ]
        return u'^BY' + self._generate_arguments(arguments, barcode_format)

    def _field_data_start(self):
        """ Add the data start command to the buffer """
        return u'^FD'

    def _field_reverse_print(self):
        """ Allows the printed data to appear white over black, or black over white
        """
        return u'^FR'

    def _field_data_stop(self):
        """ Add the data stop command to the buffer """
        return u'^FS'

    def _field_data(self, data):
        """ Add data to the buffer, between start and stop commands """
        command = u'{start}{data}{stop}'.format(
            start=self._field_data_start(),
            data=data,
            stop=self._field_data_stop(),
        )
        return command

    def font_data(self, right, down, field_format, data):
        """ Add a full text in the buffer, with needed formatting commands """
        reverse = u''
        if field_format.get(ARG_REVERSE_PRINT, False):
            reverse = self._field_reverse_print()
        block = u''
        if field_format.get(ARG_IN_BLOCK, False):
            block = self._field_block(field_format)

        command = u'{origin}{font_format}{reverse}{block}{data}'.format(
            origin=self._field_origin(right, down),
            font_format=self._font_format(field_format),
            reverse=reverse,
            block=block,
            data=self._field_data(data),
        )
        self._write_command(command)

    def barcode_data(self, right, down, barcodeType, barcode_format, data):
        """ Add a full barcode in the buffer, with needed formatting commands
        """
        command = u'{default}{origin}{barcode_format}{data}'.format(
            default=self._barcode_field_default(barcode_format),
            origin=self._field_origin(right, down),
            barcode_format=self._barcode_format(barcodeType, barcode_format),
            data=self._field_data(data),
        )
        self._write_command(command)

    def graphic_box(self, right, down, graphic_format):
        """ Send the commands to draw a rectangle """
        arguments = [
            ARG_WIDTH,
            ARG_HEIGHT,
            ARG_THICKNESS,
            ARG_COLOR,
            ARG_ROUNDING,
        ]

        # Check that the thickness value fits in the allowed values
        if graphic_format.get(ARG_THICKNESS) is not None:
            graphic_format[ARG_THICKNESS] = self._enforce(
                graphic_format[ARG_THICKNESS])

        # Check that the width value fits in the allowed values
        if graphic_format.get(ARG_WIDTH) is not None:
            graphic_format[ARG_WIDTH] = self._enforce(
                graphic_format[ARG_WIDTH],
                minimum=graphic_format[ARG_THICKNESS])

        # Check that the height value fits in the allowed values
        if graphic_format.get(ARG_HEIGHT) is not None:
            graphic_format[ARG_HEIGHT] = self._enforce(
                graphic_format[ARG_HEIGHT],
                minimum=graphic_format[ARG_THICKNESS])

        # Check that the rounding value fits in the allowed values
        if graphic_format.get(ARG_ROUNDING) is not None:
            graphic_format[ARG_ROUNDING] = self._enforce(
                graphic_format[ARG_ROUNDING], minimum=0, maximum=8)

        # Generate the ZPL II command
        command = u'{origin}{data}{stop}'.format(
            origin=self._field_origin(right, down),
            data=u'^GB' + self._generate_arguments(arguments, graphic_format),
            stop=self._field_data_stop(),
        )
        self._write_command(command)

    def graphic_circle(self, right, down, graphic_format):
        """ Send the commands to draw a circle """
        arguments = [ARG_DIAMETER, ARG_THICKNESS, ARG_COLOR]

        # Check that the diameter value fits in the allowed values
        if graphic_format.get(ARG_DIAMETER) is not None:
            graphic_format[ARG_DIAMETER] = self._enforce(
                graphic_format[ARG_DIAMETER], minimum=3, maximum=4095)

        # Check that the thickness value fits in the allowed values
        if graphic_format.get(ARG_THICKNESS) is not None:
            graphic_format[ARG_THICKNESS] = self._enforce(
                graphic_format[ARG_THICKNESS], minimum=2, maximum=4095)

        # Generate the ZPL II command
        command = u'{origin}{data}{stop}'.format(
            origin=self._field_origin(right, down),
            data=u'^GC' + self._generate_arguments(arguments, graphic_format),
            stop=self._field_data_stop(),
        )
        self._write_command(command)

    @staticmethod
    def encode_grf_ascii(pil_image):
        """Converts a PIL image to a .GRF file encoded in ascii for use
        with zebra printers.
        @:return array of strings, each string is a pixel row, containing
        4 pixel per ascii character
        """

        def round_up(value, multiple):
            result = value
            rest = value % multiple
            if rest:
                result += multiple - rest
            return result

        def get_byte_val(bits, word_len):
            if len(bits) == 0:
                return 0
            else:
                byte_value = bits[0] << word_len - 1
                return byte_value + get_byte_val(bits[1:], word_len - 1)

        # make sure we have a monochrome image, and load it.
        data = pil_image.convert("1").load()
        sizex, sizey = pil_image.size

        # increase row size to the nearest multiple of 8 pixels
        # The length of the first row of pixels is set in bytes and must be
        # a multiple of 8 bytes.
        # The zebra tools also round the x size of images to the nearest of
        # 8 pixels.
        row_size = round_up(sizex, 8)
        # construct an array y - x
        # the first row of the array, is the first line of pixels in the
        # image
        # Fill all values with white
        pixels = [[0 for x in range(0, row_size)] for y in range(0, sizey)]
        for xpos in range(0, sizex):
            for ypos in range(0, sizey):
                input_value = data[xpos, ypos]
                # PIL: pixel value 0 = white , pixel value 255 = black
                # GRF: 1=white, 0=black
                bit_value = input_value >> 7 ^ 1
                pixels[ypos][xpos] = bit_value

        ret = []
        for pixel_row in pixels:
            pixel_string = u''
            for xpos in range(0, len(pixel_row), 4):
                bits = pixel_row[xpos:xpos + 4]
                byte_value = get_byte_val(bits, 4)
                pixel_string += u'%X' % byte_value
            ret.append(pixel_string)

        return ret

    def graphic_image(self, right, down, pil_image):
        """ Print the image_data from the file handle """

        img_size_y = pil_image.size[1]
        ascii_data = self.encode_grf_ascii(pil_image)

        # The bytes per row, is the number of chars in the first row, divided
        # by 8
        # Should be equal to the X size of the image, divided by 8 and
        # rounded up
        bytes_per_row = len(ascii_data[0])/2

        # Total of bytes is the total number of pixels in the image
        # transferred to the printer divided by 8
        # Is equal to the number of bytes of the first row times the
        # number of rows
        total_bytes = bytes_per_row * img_size_y

        graphic_image_command = (
            u'^GFA,%(total_bytes)s,%(total_bytes)s,%(bytes_per_row)s,'
            u'%(ascii_data)s' % {
                u'total_bytes': total_bytes,
                u'bytes_per_row': bytes_per_row,
                u'ascii_data': u''.join(ascii_data),
            }
        )

        # Generate the ZPL II command
        command = u'{origin}{data}{stop}'.format(
            origin=self._field_origin(right, down),
            data=graphic_image_command,
            stop=self._field_data_stop(),
        )
        self._write_command(command)
