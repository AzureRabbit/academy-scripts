# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903,W0110,W0141
""" Answer class
"""

class Answer(object):
    """ Answer object
    """

    def __init__(self, answer_text, is_right=False):
        self._text = answer_text
        self._is_right = is_right

    def to_xml(self, order=0):
        """  Return an XML representation of the object

        @param order: order occupies the answer in the question
        """
        answer_pattern = (
            u'        <answer order="{}">           \n'
            u'            <text>{}</text>           \n'
            u'            <is_right>{}</is_right>   \n'
            u'        </answer>                     '
        )

        return answer_pattern.format(order, self._text, self._is_right)

