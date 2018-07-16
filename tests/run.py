import unittest
from cast.application.test import run
from cast.application import create_postgres_engine
import logging
  
logging.root.setLevel(logging.DEBUG)

class TestIntegration(unittest.TestCase):

    def test1(self):
        run(kb_name='fromjenkins83_local', application_name='AProject', engine=create_postgres_engine(user='operator',password='CastAIP',host='gaicvmdev',port=2280))  
        


if __name__ == "__main__":
    unittest.main()

