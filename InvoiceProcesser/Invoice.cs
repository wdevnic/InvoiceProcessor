using System;


namespace InvoiceProcessor
{
    /// <summary>
    /// Used to create invoice object which is comparable
    /// </summary>
    public class Invoice: IEquatable<Invoice>
    {
        // instance variables declared
        private DateTime invoiceDate;
        private string invoiceNumber;
        private string company;
        private string firstName;
        private string lastName;
        private DateTime invoicePeriodStart;
        private string location;

        private double mainRevenue;
        private double otherServices;
        private double otherServices2;
        private double revenue;
        private double productPurchases;
        private double adminFees;
        private double productPurchaseGST;
        private double adminGST;
        private double equipmentRental;
        private double benefits;

        private double netPayment;


        public Invoice()
        {
            InvoiceDate = invoiceDate;
        }

        // Invoice constructor
        public Invoice(DateTime invoiceDate, string invoiceNumber, string company, string firstName, string lastName, DateTime invoicePeriodStart, string location, double revenue, double mainRevenue, 
                        double otherServices, double otherServices2, double gstCollected, double pstCollected, double productPurchases, double adminFees, double productPurchaseGST, double adminGST, 
                       double equipmentRental, double benefits = 0)
        {
            // set invocie properties
            InvoiceDate = invoiceDate;
            InvoiceNumber = invoiceNumber;
            Company = company;
            FirstName = firstName;
            LastName = lastName;
            InvoicePeriodStart = invoicePeriodStart;
            Location = location;

            MainRevenue = mainRevenue;
            OtherServices = otherServices;
            OtherServices2 = otherServices2;
            Revenue = revenue;
            GSTCollected = gstCollected;
            PSTCollected = pstCollected;
            ProductPurchases = productPurchases;
            AdminFees = adminFees;
            ProductPurchasesGST = productPurchaseGST;
            AdminGST = adminGST;
            EquipmentRental = equipmentRental;
            Benefits = benefits;          
        }


        // setup object properties with respective exceptions thrown if invalid values submitted
        public DateTime InvoiceDate
        {
            get { return invoiceDate; }
            set
            {                    
                if (value < DateTime.Now)
                {
                    invoiceDate = value;
                }
                else
                {
                    throw new Exception("Invalid invoice date");
                }
            }
        }

        public double MainRevenue
        {
            get { return mainRevenue; }
            set
            {
                if (value >= 0.0)
                {
                    mainRevenue = value;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Invalid MainRevenue value");
                }

            }
        }

        public double OtherServices
        {
            get { return otherServices; }
            set
            {
                if (value >= 0.0)
                {
                    otherServices = value;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Invalid OtherServices value");
                }

            }
        }

        public double OtherServices2
        {
            get { return otherServices2; }
            set
            {
                if (value >= 0.0)
                {
                    otherServices2 = value;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Invalid OtherServices2 value");
                }

            }
        }


        public string InvoiceNumber
        {
            get { return invoiceNumber; }
            set
            {
                if(!String.IsNullOrEmpty(value))
                {
                    invoiceNumber = value;
                }
                else
                {
                    throw new ArgumentNullException("Invalid invoice number");
                }
            }
        }

        public string Company
        {
            get { return company; }
            set
            {
                if (!String.IsNullOrEmpty(value))
                {
                    company = value;
                }
                else
                {
                    throw new ArgumentNullException("Invalid company");
                }

            }
        }

        public string FirstName
        {
            get { return firstName; }
            set
            {
                if (!String.IsNullOrEmpty(value))
                {
                    firstName = value;
                }
                else
                {
                    throw new ArgumentNullException("Invalid first name");
                }

            }
        }

        public string LastName
        {
            get { return lastName; }
            set
            {
                if (!String.IsNullOrEmpty(value))
                {
                    lastName = value;
                }
                else
                {
                     throw new ArgumentNullException("Invalid last name");
                }
            }
        }

        public DateTime InvoicePeriodStart
        {
            get { return invoicePeriodStart; }
            set
            {
                if(value < DateTime.Now)
                {
                    invoicePeriodStart = value;                
                }
                else
                {
                    throw new Exception("Invalid invoice period start date");
                }
            }
        }

        public string Location
        {
            get { return location; }
            set
            {
                if (!String.IsNullOrEmpty(value))
                {
                    location = value;
                }
                else
                {
                    throw new ArgumentNullException("Invalid location");
                }
            }
        }

        public double Revenue
        {
            get { return revenue; }
            set
            {
                if(value >= 0.0)
                {
                    revenue = value;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Invalid revenue value");
                }

            }
        }

        public double GSTCollected { get; set; }

        public double PSTCollected { get; set; }

        public double ProductPurchases
        {
            get { return productPurchases; }
            set
            {
                if (value >= 0.0)
                {
                    productPurchases = value;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Invalid product purchases value");
                }
            }
        }

        public double AdminFees
        {
            get { return adminFees; }
            set
            {
                if (value >= 0.0)
                {
                    adminFees = value;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Invalid admin fees value");
                }
            }
        }

        public double ProductPurchasesGST
        {
            get { return productPurchaseGST; }
            set
            {
                if (value >= 0.0)
                {
                    productPurchaseGST = value;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Invalid product purchase PST value");
                }
            }
        }

        public double AdminGST
        {
            get { return adminGST; }
            set
            {
                if (value >= 0.0)
                {
                    adminGST = value;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Invalid admin GST value");
                }
            }
        }

        public double EquipmentRental
        {
            get { return equipmentRental; }
            set
            {
                if (value >= 0.0)
                {
                    equipmentRental = value;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Invalid equipment rental value");
                }
            }
        }

        public double Benefits
        {
            get { return benefits; }
            set
            {
                if (value >= 0.0)
                {
                    benefits = value;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Invalid benefits value");
                }
            }
        }

        public double NetPayment
        {
            get { return netPayment; }
            set
            {
                netPayment = (Revenue + GSTCollected + PSTCollected) - (ProductPurchases + AdminFees + ProductPurchasesGST + AdminGST + EquipmentRental);
            }
             
        }

        /// <summary>
        /// Implement Equal method
        /// </summary>
        /// <param name="other">Another invoice</param>
        /// <returns>Whether 2 invocies are equal based on invoice number</returns>
        public bool Equals(Invoice other)
        {
            if(other == null)
            {
                return false;
            }

            if(this.InvoiceNumber == other.InvoiceNumber)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        // tostring override
        public override string ToString()
        {
            return "Invoice: " + InvoiceDate + " " + InvoiceNumber + " " + Company + " " + FirstName + " " + LastName + " " + InvoicePeriodStart + " " + Location + " " +
                Revenue + " " + MainRevenue + " " + OtherServices + " " + OtherServices2 + " " + GSTCollected + " " + PSTCollected + " " + ProductPurchases + " " + AdminFees + " " + ProductPurchasesGST + " " + AdminGST + " " +
                EquipmentRental + " " + Benefits ;
        }
    } 
      
}
