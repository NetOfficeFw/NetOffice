using System;
using System.Collections.Generic;
using System.Text;

namespace COMAddinTaskPaneExampleCS4
{
    class Customer
    {
        int _id;
        string _name;
        string _company;
        string _city;
        string _postalCode;
        string _country;
        string _phone;

        internal Customer(int id, string name, string company, string city, string postalCode, string country, string phone)
        {
            _id = id;
            _name = name;
            _company = company;
            _city = city;
            _postalCode = postalCode;
            _country = country;
            _phone = phone;
        }

        public int ID { get { return _id; } }
        public string Name { get { return _name; } }
        public string Company { get { return _company; } }
        public string City { get { return _city; } }
        public string PostalCode { get { return _postalCode; } }
        public string Country { get { return _country; } }
        public string Phone { get { return _phone; } }


        public override string ToString()
        {
            return string.Format("Name: {1}{7}ID: {0}{7}Company: {2}{7}City: {3}{7}PostalCode: {4}{7}Country: {5}{7}Phone: {6}{7}{7}", _id, _name, _company, _city, _postalCode, _country, _phone, Environment.NewLine);
        }
    }
}
