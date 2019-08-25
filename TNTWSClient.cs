using System;
using System.Collections.Generic;
using System.Text;
using TNTWSSample.TNTWebReference;
using System.ServiceModel;
using System.Data.OleDb;
using System.Globalization;

namespace TNTWSSample
{
    class TNTWSClient
    {
       
        static void Main()
        {
            // Connection string and SQL query
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=K:\POS Database\POS Construction_be.mdb";
            string strSQL = "SELECT ID, si_SaturdayDeliveryOption, p_TrackingNumber, si_ServiceType, UPSTNTDay, si_ReturnServiceOption, si_PickUpDate, st_City, st_State, st_PostalZipCode, st_Country, RecipientCode FROM TRACKING WHERE((UPSTNTDay) = '' or (isNull(UPSTNTDay) AND ((si_ReturnServiceOption) = 'N') AND ((si_VoidIndicator) = 'N')))";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Create a command and set its connection  
                OleDbCommand command = new OleDbCommand(strSQL, connection);
                try
                {
                    connection.Open();
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                string strTrackID = reader.GetString(reader.GetOrdinal("p_TrackingNumber"));
                                string strSatDelOpt = reader.GetString(reader.GetOrdinal("si_SaturdayDeliveryOption"));
                                string strCity = reader.GetString(reader.GetOrdinal("st_City"));
                                string strState = reader.GetString(reader.GetOrdinal("st_State"));
                                string strZip = reader.GetString(reader.GetOrdinal("st_PostalZipCode"));
                                string strShipdate = reader.GetString(reader.GetOrdinal("si_PickUpDate"));
                                string strCountry = reader.GetString(reader.GetOrdinal("st_Country"));
                                string strServiceType = reader.GetString(reader.GetOrdinal("si_ServiceType"));
                                Int32 strID = reader.GetInt32(reader.GetOrdinal("ID"));
                                
                                try
                                {
                                    TimeInTransitService tntService = new TimeInTransitService();
                                    TimeInTransitRequest tntRequest = new TimeInTransitRequest();
                                    RequestType request = new RequestType();
                                    String[] requestOption = { "TNT" };
                                    request.RequestOption = requestOption;
                                    tntRequest.Request = request;
                                    tntRequest.SaturdayDeliveryInfoRequestIndicator = "N";
                                    RequestShipFromType shipFrom = new RequestShipFromType();
                                    RequestShipFromAddressType addressFrom = new RequestShipFromAddressType();
                                    addressFrom.City = "Edina";
                                    addressFrom.CountryCode = "US";
                                    addressFrom.PostalCode = "55435";
                                    addressFrom.StateProvinceCode = "MN";
                                    shipFrom.Address = addressFrom;
                                    tntRequest.ShipFrom = shipFrom;
                                    RequestShipToType shipTo = new RequestShipToType();
                                    RequestShipToAddressType addressTo = new RequestShipToAddressType();
                                    addressTo.City = strCity;
                                    addressTo.CountryCode = strCountry;
                                    addressTo.PostalCode = strZip;
                                    addressTo.StateProvinceCode = strState;
                                    shipTo.Address = addressTo;
                                    tntRequest.ShipTo = shipTo;
                                    PickupType pickup = new PickupType();
                                    string left = strShipdate.Substring(0, 8);
                                    pickup.Date = left;
                                    pickup.Time = "120000";
                                    tntRequest.Pickup = pickup;

                                    if (strCountry == "CA")
                                    {
                                        ShipmentWeightType shipmentWeight = new ShipmentWeightType();
                                        shipmentWeight.Weight = "10";
                                        CodeDescriptionType unitOfMeasurement = new CodeDescriptionType();
                                        unitOfMeasurement.Code = "KGS";
                                        unitOfMeasurement.Description = "Kilograms";
                                        shipmentWeight.UnitOfMeasurement = unitOfMeasurement;
                                        tntRequest.ShipmentWeight = shipmentWeight;
                                        tntRequest.TotalPackagesInShipment = "1";
                                        InvoiceLineTotalType invoiceLineTotal = new InvoiceLineTotalType();
                                        invoiceLineTotal.CurrencyCode = "CAD";
                                        invoiceLineTotal.MonetaryValue = "10";
                                        tntRequest.InvoiceLineTotal = invoiceLineTotal;
                                        tntRequest.MaximumListSize = "1";
                                    }
                                    else
                                    {
                                        ShipmentWeightType shipmentWeight = new ShipmentWeightType();
                                        shipmentWeight.Weight = "10";
                                        CodeDescriptionType unitOfMeasurement = new CodeDescriptionType();
                                        unitOfMeasurement.Code = "LBS";
                                        unitOfMeasurement.Description = "pounds";
                                        shipmentWeight.UnitOfMeasurement = unitOfMeasurement;
                                        tntRequest.ShipmentWeight = shipmentWeight;
                                        tntRequest.TotalPackagesInShipment = "1";
                                        InvoiceLineTotalType invoiceLineTotal = new InvoiceLineTotalType();
                                        invoiceLineTotal.CurrencyCode = "USD";
                                        invoiceLineTotal.MonetaryValue = "10";
                                        tntRequest.InvoiceLineTotal = invoiceLineTotal;
                                        tntRequest.MaximumListSize = "1";
                                    }

                                    UPSSecurity upss = new UPSSecurity();
                                    UPSSecurityServiceAccessToken upsSvcToken = new UPSSecurityServiceAccessToken();
                                    upsSvcToken.AccessLicenseNumber = "1D5E2960D39CB1B5 ";
                                    upss.ServiceAccessToken = upsSvcToken;
                                    UPSSecurityUsernameToken upsSecUsrnameToken = new UPSSecurityUsernameToken();
                                    upsSecUsrnameToken.Username = "chad jasicki";
                                    upsSecUsrnameToken.Password = "asdf46#$2";
                                    upss.UsernameToken = upsSecUsrnameToken;
                                    tntService.UPSSecurityValue = upss;
                                   
                                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12 | System.Net.SecurityProtocolType.Tls | System.Net.SecurityProtocolType.Tls11; //This line will ensure the latest security protocol for consuming the web service call.
                                    TimeInTransitResponse tntResponse = tntService.ProcessTimeInTransit(tntRequest);

                                    if(strServiceType == "WORLDWIDE SAVER") {strServiceType = "Worldwide Saver";}

                                    strServiceType = "UPS " + strServiceType;

                                    if (tntResponse.Item != null)
                                    {
                                        var timeInTransitResponse = (TransitResponseType)tntResponse.Item;
                                        foreach (var serviceSummaryType in timeInTransitResponse.ServiceSummary)
                                        {
                                            string strUPScode = serviceSummaryType.Service.Code;

                                            if (serviceSummaryType.Service.Description == strServiceType && strSatDelOpt == "Y" && strUPScode.Substring(strUPScode.Length - 1, 1) == "S" || serviceSummaryType.Service.Description == strServiceType && strSatDelOpt != "Y" && strUPScode.Substring(strUPScode.Length - 1, 1) != "S")
                                            {
                                                        Console.WriteLine(addressTo.City + strState + ", " + strCity + " - " + serviceSummaryType.EstimatedArrival.BusinessDaysInTransit);
                                                        string intUPSTNT = serviceSummaryType.EstimatedArrival.BusinessDaysInTransit;
                                                        string strTNTDay = serviceSummaryType.EstimatedArrival.Arrival.Date;
                                                        DateTime d = DateTime.ParseExact(strTNTDay, "yyyyMMdd", CultureInfo.InvariantCulture);
                                                        string strTNTtime = serviceSummaryType.EstimatedArrival.Arrival.Time;

                                                        if (Convert.ToInt32(strTNTtime) >= 230000)
                                                        {
                                                            strTNTtime = "End of Day";
                                                        }
                                                        else
                                                        {
                                                            IFormatProvider format = System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat;
                                                            DateTime time_24 = DateTime.ParseExact(strTNTtime, "HHmmss", format);
                                                            strTNTtime = time_24.ToString("h:mm tt");

                                                        }

                                                        strSQL = "UPDATE TRACKING SET UPSTNTDay = '" + d.ToString("MM/dd/yyyy") + "', UPSTNTTime = '" + strTNTtime + "' WHERE p_TrackingNumber ='" + strTrackID + "'";
                                                        command = new OleDbCommand(strSQL, connection);
                                                        command.ExecuteReader();

                                                        Console.Write(strTrackID);
                                                        Console.Write("Business Days in Transit: ");
                                                        Console.Write(serviceSummaryType.EstimatedArrival.BusinessDaysInTransit);
                                                        Console.Write(", Arrival Date: ");
                                                        Console.Write(serviceSummaryType.EstimatedArrival.Arrival.Date);
                                                        Console.Write(", Service: (");
                                                        Console.Write(serviceSummaryType.Service.Code);
                                                        Console.Write(") ");
                                                        Console.Write(serviceSummaryType.Service.Description);
                                                        Console.WriteLine(serviceSummaryType.EstimatedArrival.Arrival.Time);
                                            }
                                        }
                                    }   
                                }
                                catch (System.Web.Services.Protocols.SoapException ex)
                                {
                                    Console.WriteLine("");
                                    Console.WriteLine("---------Time In Transit Web Service returns error----------------");
                                    Console.WriteLine("---------\"Hard\" is user error \"Transient\" is system error----------------");
                                    Console.WriteLine("SoapException Message= " + ex.Message);
                                    Console.WriteLine("");
                                    Console.WriteLine("SoapException Category:Code:Message= " + ex.Detail.LastChild.InnerText);
                                    Console.WriteLine("");
                                    Console.WriteLine("SoapException XML String for all= " + ex.Detail.LastChild.OuterXml);
                                    Console.WriteLine("");
                                    Console.WriteLine("SoapException StackTrace= " + ex.StackTrace);
                                    Console.WriteLine("-------------------------");
                                    Console.WriteLine("");
                                }
                                catch (System.ServiceModel.CommunicationException ex)
                                {
                                    Console.WriteLine("");
                                    Console.WriteLine("--------------------");
                                    Console.WriteLine("CommunicationException= " + ex.Message);
                                    Console.WriteLine("CommunicationException-StackTrace= " + ex.StackTrace);
                                    Console.WriteLine("-------------------------");
                                    Console.WriteLine("");

                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("");
                                    Console.WriteLine("-------------------------");
                                    Console.WriteLine(" Generaal Exception= " + ex.Message);
                                    Console.WriteLine(" Generaal Exception-StackTrace= " + ex.StackTrace);
                                    Console.WriteLine("-------------------------");

                                }
                                finally
                                {
                                }
                            }
                        }
                    }
                }
                catch (System.ServiceModel.CommunicationException ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.WriteLine(" General Exception-StackTrace= " + ex.StackTrace);
                }
            }
            Console.WriteLine("done");
            //Console.ReadKey();
        }
    }
}