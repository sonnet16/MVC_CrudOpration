using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using StudentInfoSystem.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace StudentInfoSystem.Controllers
{
    public class HomeController : Controller
    {
        StudentDBOperation stdb = new StudentDBOperation();
        public ActionResult Index()
        {
            return View();
        }

        public JsonResult List()
        {
            return Json(stdb.ListAll(), JsonRequestBehavior.AllowGet);
        }
        public JsonResult Add(StudentInfo student)
        {
            return Json(stdb.AddProduct(student), JsonRequestBehavior.AllowGet);
        }

        public ActionResult WorkOrderPrintList(string sIDs)
        {
            StudentInfo _oWorkOrder = new StudentInfo();
           List<StudentInfo> _oWorkOrders = new List<StudentInfo>();
            //string sSql = "SELECT * FROM View_WorkOrder WHERE WorkOrderID IN (" + sIDs + ")";
            //_oWorkOrders = WorkOrder.Gets(sSql, (int)Session[SessionInfo.currentUserID]);
            //int nUserID = Convert.ToInt32(Session[SessionInfo.currentUserID]);
            //_oWorkOrder.WorkOrderList = _oWorkOrders;
            _oWorkOrders = stdb.ListAll();
            if (_oWorkOrders.Count > 0)
            {
                //Company oCompany = new Company();
                //oCompany = oCompany.Get(1, (int)Session[SessionInfo.currentUserID]);
                //oCompany.CompanyLogo = GetCompanyLogo(oCompany);
                //BusinessUnit oBusinessUnit = new BusinessUnit();
                //oBusinessUnit = oBusinessUnit.Get(_oWorkOrders[0].BUID, (int)Session[SessionInfo.currentUserID]);
                //oCompany.Name = oBusinessUnit.Name;
                //oCompany.Address = oBusinessUnit.Address;
                //oCompany.Phone = oBusinessUnit.Phone;
                //oCompany.Email = oBusinessUnit.Email;
                //oCompany.WebAddress = oBusinessUnit.WebAddress;
                //_oWorkOrder.Company = oCompany;

                rptWorkOrderList oReport = new rptWorkOrderList();
                byte[] abytes = oReport.PrepareReport(_oWorkOrders);
                return File(abytes, "application/pdf");
            }
            else
            {

                string sMessage = "There is no data for print";
                return RedirectToAction("MessageHelper", "User", new { message = sMessage });
            }

        }

        public void ExportXL()
        {
            StudentInfo oUser = new StudentInfo();
            List<StudentInfo> oUsers = new List<StudentInfo>();
            //try
            //{
            //    oUser = (ESimSol.BusinessObjects.User)Session[SessionInfo.ParamObj];
            //    if (oUser.LogInID == "")
            //    {
            //        oUser.LogInID = "0";
            //    }
            //    string sSQL = "SELECT * FROM View_User AS HH WHERE HH.UserID IN (" + oUser.LogInID + ") ORDER BY HH.UserName ASC";
            //    oUsers = ESimSol.BusinessObjects.User.GetsBySql(sSQL, (int)Session[SessionInfo.currentUserID]);
            //}
            //catch (Exception ex)
            //{
            //    oUsers = new List<ESimSol.BusinessObjects.User>();
            //}
            oUsers = stdb.ListAll();
            if (oUsers.Count > 0)
            {
                //Company oCompany = new Company();
                //oCompany = oCompany.Get(1, ((User)Session[SessionInfo.CurrentUser]).UserID);

                #region Header
                List<TableHeader> table_header = new List<TableHeader>();
                table_header.Add(new TableHeader { Header = "#SL", Width = 10f, IsRotate = false });
                table_header.Add(new TableHeader { Header = "Student Name", Width = 25f, IsRotate = false });
                table_header.Add(new TableHeader { Header = "DateOfBirth", Width = 15f, IsRotate = false });
                table_header.Add(new TableHeader { Header = "BloodGroup", Width = 35f, IsRotate = false });
                table_header.Add(new TableHeader { Header = "MaritalStatus", Width = 20f, IsRotate = false });
                table_header.Add(new TableHeader { Header = "Gender", Width = 20f, IsRotate = false });
                table_header.Add(new TableHeader { Header = "Reliogion", Width = 30f, IsRotate = false });

                #endregion

                #region Export Excel
                int nRowIndex = 1, nStartCol = 1, nEndCol = table_header.Count;
                ExcelRange cell; ExcelFill fill;
                OfficeOpenXml.Style.Border border;

                using (var excelPackage = new ExcelPackage())
                {
                    excelPackage.Workbook.Properties.Author = "ESimSol";
                    excelPackage.Workbook.Properties.Title = "Export from ESimSol";
                    var sheet = excelPackage.Workbook.Worksheets.Add("StudentInfo");
                    sheet.View.FreezePanes(5, 4);

                    foreach (TableHeader listItem in table_header)
                    {
                        sheet.Column(nStartCol++).Width = listItem.Width;
                    }

                    //#region Report Header
                    //cell = sheet.Cells[nRowIndex, 1, nRowIndex, nEndCol + 1]; cell.Merge = true; cell.Value = oCompany.Name; cell.Style.Font.Bold = true;
                    //cell.Style.Font.Size = 20; cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //nRowIndex++;

                    //cell = sheet.Cells[nRowIndex, 1, nRowIndex, nEndCol + 1]; cell.Merge = true; cell.Value = oCompany.Address; cell.Style.Font.Bold = true;
                    //cell.Style.Font.Size = 12; cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //nRowIndex++;

                    //cell = sheet.Cells[nRowIndex, 1, nRowIndex, nEndCol + 1]; cell.Merge = true; cell.Value = "ESimSol User List"; cell.Style.Font.Bold = true;
                    //cell.Style.Font.Size = 12; cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //nRowIndex += 1;
                    #endregion

                    #region Column Header
                    nStartCol = 1;
                    foreach (TableHeader listItem in table_header)
                    {
                        cell = sheet.Cells[nRowIndex, nStartCol++]; cell.Value = listItem.Header; cell.Style.Font.Bold = true; cell.Style.WrapText = true;
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
                    }
                    nRowIndex++;
                    #endregion


                    #region Data
                    int nSL = 0;
                    foreach (StudentInfo oItem in oUsers)
                    {
                        nStartCol = 1;
                        ExcelTool.FillCellBasic(sheet, nRowIndex, nStartCol++, (++nSL).ToString(), false, ExcelHorizontalAlignment.Center, false, false);
                        ExcelTool.FillCellBasic(sheet, nRowIndex, nStartCol++, oItem.StudentName, false, ExcelHorizontalAlignment.Left, false, false);
                        ExcelTool.FillCellBasic(sheet, nRowIndex, nStartCol++, oItem.DateOfBirthSt, false, ExcelHorizontalAlignment.Left, false, false);
                        ExcelTool.FillCellBasic(sheet, nRowIndex, nStartCol++, oItem.BloodGroup, false, ExcelHorizontalAlignment.Left, false, false);
                        ExcelTool.FillCellBasic(sheet, nRowIndex, nStartCol++, oItem.MaritalStatus, false, ExcelHorizontalAlignment.Left, false, false);
                        ExcelTool.FillCellBasic(sheet, nRowIndex, nStartCol++, oItem.Gender, false, ExcelHorizontalAlignment.Left, false, false);    
                        ExcelTool.FillCellBasic(sheet, nRowIndex, nStartCol++, oItem.Reliogion, false, ExcelHorizontalAlignment.Center, false, false);
                        nRowIndex++;

                    }
                    #endregion



                    cell = sheet.Cells[1, 1, nRowIndex, table_header.Count + 2];
                    fill = cell.Style.Fill; fill.PatternType = ExcelFillStyle.Solid;
                    fill.BackgroundColor.SetColor(Color.White);

                    Response.ClearContent();
                    Response.BinaryWrite(excelPackage.GetAsByteArray());
                    Response.AddHeader("content-disposition", "attachment; filename=StudentInfo.xlsx");
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.Flush();
                    Response.End();
                }
          
            }
        }
    }
}