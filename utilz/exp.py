'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-06-12
@Last Modified: 2019-06-12 8:26:56 am
@Modified by:   zmFeng
expression builder
'''

from utilz import triml

class AbsResolver(object):
    ''' a class to resolve argument passed to it, this default one return
    any argument passed to it
    '''
    def resolve(self, arg):
        ''' resolve the given arg to value
        '''
        return arg

_df_rsv = AbsResolver()

class Exp(object):
    '''
    an expression that should be able to evaluate itself
    '''
    def __init__(self, lopr, op, ropr=None):
        '''
        Args:
            lopr: the left operator
            op: the operation, can be one of +,-,*,/,and,or,not,==,>,>=,<,<=,!=
            ropr: the right operator
        '''
        self._op, self._lopr, self._ropr = op, lopr, ropr

    def chain(self, op, *oprs):
        ''' return a chain expression, an often used example is
        a.chain('and', expb, expc, expf)
        '''
        e = self
        for opr in oprs:
            e = Exp(e, op, opr)
        return e

    def and_(self, exp):
        ''' convenient "and" operation
        '''
        return Exp(self, 'and', exp)

    def or_(self, exp):
        ''' convenient "or" operation
        '''
        return Exp(self, 'or', exp)

    def eval(self, resolver=None):
        '''
        evaluate myself
        Args:
            resolver: an instance that sub-class AbsResolver, who can solve the operators
        '''
        if not resolver:
            resolver = _df_rsv
        l = self._lopr.eval(resolver) if isinstance(self._lopr, Exp) else resolver.resolve(self._lopr)
        r = self._ropr.eval(resolver) if isinstance(self._ropr, Exp) else resolver.resolve(self._ropr)
        op, rc = self._op, None
        if isinstance(op, str):
            op = triml(op)
            if op == 'and':
                rc = l and r
            elif op == 'or':
                rc = l or r
            elif op == 'not':
                rc = not l
            elif op == '==':
                rc = l == r
            elif op == '!=':
                rc = l != r
            elif op == '>':
                rc = l > r
            elif op == '>=':
                rc = l >= r
            elif op == '<':
                rc = l < r
            elif op == '<=':
                rc = l <= r
            elif op in ('+', 'add'):
                rc = l + r
            elif op == '-':
                rc = l - r
            elif op == '*':
                rc = l * r
            elif op == '/':
                rc = l / r
            else:
                rc = self._eval_ext(op, l, r)
        else:
            rc = op(l)
        return rc

    def __str__(self):
        return '(%s %s %s)' % (self._lopr, self._op, self._ropr)

    def _eval_ext(self, op, l, r):
        ''' when the basic function eval can not process, this will be called,
        extend this in the sub-class to do more op
        '''
        return None
