# coding=utf-8 
"""
 * @Author: zmFeng 
 * @Date: 2018-05-24 16:21:19 
 * @Last Modified by:   zmFeng 
 * @Last Modified time: 2018-05-24 16:21:19 
 * examples from sqlalchemy
 """
from sqlalchemy.sql.sqltypes import Integer,String
from sqlalchemy.sql.schema import Column,ForeignKey
from sqlalchemy.orm import relationship
from sqlalchemy.ext.declarative.api import declarative_base

Base = declarative_base()

class Address(Base):
     __tablename__ = 'addresses'
     id = Column(Integer, primary_key=True)
     email_address = Column(String, nullable=False)
     user_id = Column(Integer, ForeignKey('users.id'))

     user = relationship("User", back_populates="addresses")

     def __repr__(self):
         return "<Address(email_address='%s')>" % self.email_address

class User(Base):
     __tablename__ = 'users'

     id = Column(Integer, primary_key=True)
     name = Column(String)
     fullname = Column(String)
     password = Column(String)
     addresses = relationship("Address", order_by=Address.id, back_populates="user")
     def __repr__(self):
        return "<User(name='%s', fullname='%s', password='%s')>" % (
                             self.name, self.fullname, self.password)