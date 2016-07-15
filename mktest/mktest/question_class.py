# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903,W0110,W0141,W0403
""" Question class
"""

from answer_class import Answer

class Question(object):
    """ Question object
    """

    def __init__(self, question_text):
        self._text = question_text
        self._answers = []

    def add_answer(self, answer_text, is_right=False):
        """ Adds new answer for question

        @param answer_text: text for the answer
        @param is_right: marks answer as the right answer
        """
        answer = Answer(answer_text, is_right)
        self._answers.append(answer)

    def to_xml(self, order=0):
        """  Return an XML representation of the object

        @param order: order occupies the answer in the question
        """

        index = 0
        answers_text = u''
        for answer in self._answers:
            answers_text = answers_text + answer.to_xml(index) + u'\n'
            index = index + 1

        question_pattern = (
            u'<question order="{}"> \n'
            u'    <text>{}</text>     \n'
            u'    <answers>           \n{}'
            u'    </answers>          \n'
            u'</question>             '
        )
        return question_pattern.format(order, self._text, answers_text)
