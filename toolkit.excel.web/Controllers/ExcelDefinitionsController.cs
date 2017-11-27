using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using toolkit.excel.data;

namespace toolkit.excel.web.Controllers
{
    public class ExcelDefinitionsController : Controller
    {
        private ExcelDataContext db = new ExcelDataContext();

        // GET: ExcelDefinitions
        public ActionResult Index()
        {
            return View(db.ExcelDefinition.ToList());
        }

        // GET: ExcelDefinitions/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ExcelDefinition excelDefinition = db.ExcelDefinition.Find(id);
            if (excelDefinition == null)
            {
                return HttpNotFound();
            }
            return View(excelDefinition);
        }

        // GET: ExcelDefinitions/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: ExcelDefinitions/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "DefinitionId,FileName,SheetName,Range,TargetTable,ConnectionString,HasHeaderRow")] ExcelDefinition excelDefinition)
        {
            if (ModelState.IsValid)
            {
                db.ExcelDefinition.Add(excelDefinition);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(excelDefinition);
        }

        // GET: ExcelDefinitions/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ExcelDefinition excelDefinition = db.ExcelDefinition.Find(id);
            if (excelDefinition == null)
            {
                return HttpNotFound();
            }
            return View(excelDefinition);
        }

        // POST: ExcelDefinitions/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "DefinitionId,FileName,SheetName,Range,TargetTable,ConnectionString,HasHeaderRow")] ExcelDefinition excelDefinition)
        {
            if (ModelState.IsValid)
            {
                db.Entry(excelDefinition).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(excelDefinition);
        }

        // GET: ExcelDefinitions/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ExcelDefinition excelDefinition = db.ExcelDefinition.Find(id);
            if (excelDefinition == null)
            {
                return HttpNotFound();
            }
            return View(excelDefinition);
        }

        // POST: ExcelDefinitions/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            ExcelDefinition excelDefinition = db.ExcelDefinition.Find(id);
            db.ExcelDefinition.Remove(excelDefinition);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
