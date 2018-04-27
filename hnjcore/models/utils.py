'''
Created on Apr 17, 2018

@author: zmFeng
'''
class JOElement(object):
    """ 
    representation of Alpha + digit composite key
    the constructor method can be one of:
    
    JOElement("A1234BC")
    
    JOElement("A",123,"BC")
    
    JOElement(alpha = "A",digit = 123,suffix = "BC")
    """    
    __minlen__ = 5
    def __init__(self, *args, **kwargs):
        cnt = len(args)
        if(cnt == 1):
            self._parse_(args[0])
        elif(cnt >= 2):
            self.alpha = args[0].strip()
            self.digit = args[1]
            self.suffix = args[2].strip() if(cnt > 2) else "" 
        else:
            self._reset_()
    
    def _parse_(self,jono):
        stg = 0
        strs = ["","",""]
        jono = jono.strip()
        for i in range(len(jono)):
            if(jono[i].isalpha()):
                if(stg == 0):
                    strs[0] = strs[0] + jono[i]
                else:
                    strs[2] = strs[2] + jono[i:]
                    break 
            elif(jono[i].isdigit()):
                if(not stg):                    
                    stg += 1
                    #first character is number, let it be alpha
                    if(len(strs[0]) == 0):
                        strs[0] = jono[i]
                        continue
                strs[1] = strs[1] + jono[i]
            else:
                break
        if(stg and strs[1].isdigit()):
            self.alpha = strs[0].strip()
            self.digit = int(strs[1])
            self.suffix = strs[2].strip()
        else:
            self._reset_()
    
    def _reset_(self):
        self.alpha = ""
        self.digit = 0
        self.suffix = "" 
            
    def __repr__(self, *args, **kwargs):
        return "JOElement(%s,%d,%s)" % (self.alpha,self.digit,self.suffix)
    
    def __str__(self, *args, **kwargs):
        if(hasattr(self,'digit')):
            return self.alpha + \
                (("%0" + str(self.__minlen__ - len(self.alpha)) + "d") % self.digit) + self.suffix 
        else:
            return ""
        
    @property
    def value(self):
        return self.__str__()
    
    def __composite_values__(self):
        return self.alpha,self.digit        
    
    def __eq__(self,other):
        return isinstance(other,JOElement) and \
            self.alpha == other.alpha and \
            self.digit == other.digit
    def __ne__(self,other):
        return not self.__eq__(other)
    
    def __ge__(self,other):
        return isinstance(other,JOElement) and \
            self.alpha == other.digit and \
            self.digit >= other.digit

class StyElement(JOElement):
    def __composite_values__(self):
        pr = JOElement.__composite_values__(self)
        return pr[0],pr[1],self.suffix
    
    def __eq__(self, other):
        return JOElement.__eq__(self, other) and self.suffix == other.suffix