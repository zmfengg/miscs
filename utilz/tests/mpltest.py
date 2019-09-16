'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-09-16
@Last Modified: 2019-09-16 8:14:02 am
@Modified by:   zmFeng

Matplotlib usages
'''

from os import remove
from tempfile import TemporaryFile
from unittest import TestCase

import numpy as np
from matplotlib import pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


class MplSuite(TestCase):
    ''' matplotlib tests
    '''

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls._show_plt = True
    
    def _show(self):
        if self._show_plt:
            plt.show()
        plt.close()

    def testQuickStart(self):
        ''' one axes set only
        '''
        ax = plt.subplot()
        t1 = np.arange(0.0, 2.0, 0.1)
        t2 = np.arange(0.0, 2.0, 0.01)
        l1, = ax.plot(t2, np.exp(-t2))
        l2, l3 = ax.plot(t2, np.sin(2 * np.pi * t2), '--o', t1, np.log(1 + t1), '.')
        l4, = ax.plot(t2, np.exp(-t2) * np.sin(2 * np.pi * t2), 's-.')
        l5, = ax.plot((0, 0.1, 0.3, 0.5, 0.1), label='a line')
        ax.legend((l2, l4, l5), ('oscillatory', 'damped', 'tuple'), loc='upper right', shadow=True)
        ax.set_xlabel('time')
        ax.set_ylabel('volts')
        ax.set_title('Damped oscillation')
        self._show()
    
    def testSubplotWithId(self):
        ''' invoke several subplot(), see the output
        '''
        # calling subplot() several times will have only one axes on top, so
        # need to do more
        # 2x2 can hold up to 4, so 224 is the upper limit
        for i in range(221, 225):
            ax = plt.subplot(i)
            print(id(ax))
            ax.plot([1, 2, 3])
            ax.set_xlabel('%d of' % i)
        with self.assertRaises(ValueError):
            ax = plt.subplot(226)
        self._show()

    def testsubplots(self):
        ''' 2x2 subplots, flat, constrained_layout
        '''
        fig, axs = plt.subplots(2, 2, constrained_layout=True)
        idx = 0
        # don't use "for ax in axs" because it return narray only, flat do it
        for ax in axs.flat:
            if idx == 2:
                ax.plot((1, 2, 3, 4, 5), (1, 2, 5, 3, 2), '--x', label='Line%d' % idx)
                ln = ax.plot((1, 2, 5, 3, 2), '--^', label="won't be shown")
                ax.legend((ln[0], ), ('use ln[0], not ln',))
                ax.set_xlabel('sequence') # or plt.xlabel()
                ax.set_ylabel('value') # or plt.ylabel()
                ax.set_title('%d axes' % idx)
            else:
                ax.plot((1, 2, 5, 3, 2), '--o', label='Line%d' % idx)
                ax.legend()
            ax.set_title("%d axes" % idx)
            idx += 1
        fig.suptitle('%d-row X %d-col figure' % axs.shape)
        self._show()
    
    def testMinorGridlines(self):
        fig = plt.figure(1)
        for idx in range(221, 225):
            ax = plt.subplot(idx)
            # below 2 lines can be before or after the grid commands
            for i in range(1, 30, 5):
                x = np.linspace(0, 10, 50)
                plt.plot(x, np.sin(x) * i, label='line %d' % i)
            # grids
            if idx == 221:
                ax.grid(b=True, which='major', color='k', linestyle='-.', alpha=0.8)
                ax.grid(b=True, which='minor', color='r')
                ax.legend()
            else:
                ax.grid(b=True, which='both', linestyle='-.')
            ax.minorticks_on() # must be call for each axes
        plt.show()

        
    def testMultiPage(self):
        ''' create a 2-page 2X2 axes pdf file
        '''
        fn = TemporaryFile().name + '.pdf'
        with PdfPages(fn) as pdf:
            for cnt in range(2):
                fig, axs = plt.subplots(2, 2)
                fig.suptitle('page %d' % (cnt + 1))
                for ax in axs.flat:
                    for i in range(1, 30, 5):
                        x = np.linspace(0 + i, 10 + i, 50)
                        ax.plot(x, np.sin(x) * i, label='line %d' % i)
                    ax.grid(b=True, which='both', linestyle='-.')
                    ax.minorticks_on() # must be call for each axes
                pdf.savefig(fig)
        remove(fn)
