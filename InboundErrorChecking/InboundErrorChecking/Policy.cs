using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InboundErrorChecking
{


    class Policy
    {
        private string policy_reference;
        private string eventtype;
        private string root;
        private string errorsummary;

        public Policy(string pol, string eventt, string roott, string error)
        {
            policy_reference = pol;
            eventtype = eventt;
            root = roott;
            errorsummary = error;
        }


        public string Policyref
        {
            get
            {
                return policy_reference;
            }

            set
            {
                policy_reference = value;
            }
        }

        public string Eventtype
        {
            get
            {
                return eventtype;
            }

            set
            {
                eventtype = value;
            }
        }


        public string Root
        {
            get
            {
                return root;
            }

            set
            {
                root = value;
            }
        }

        public string Error
        {
            get
            {
                return errorsummary;
            }

            set
            {
                errorsummary = value;
            }
        }


    }


}
