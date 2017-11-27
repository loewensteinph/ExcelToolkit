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
    public class ColumnMappingsController : Controller
    {
        private ExcelDataContext db = new ExcelDataContext();

        // GET: ColumnMappings
        public ActionResult Index()
        {
            return View(db.ColumnMapping.ToList());
        }

        // GET: ColumnMappings/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ColumnMapping columnMapping = db.ColumnMapping.Find(id);
            if (columnMapping == null)
            {
                return HttpNotFound();
            }
            return View(columnMapping);
        }

        // GET: ColumnMappings/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: ColumnMappings/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ColumnMappingId,SourceColumn,TargetColumn")] ColumnMapping columnMapping)
        {
            if (ModelState.IsValid)
            {
                db.ColumnMapping.Add(columnMapping);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(columnMapping);
        }

        // GET: ColumnMappings/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ColumnMapping columnMapping = db.ColumnMapping.Find(id);
            if (columnMapping == null)
            {
                return HttpNotFound();
            }
            return View(columnMapping);
        }

        // POST: ColumnMappings/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ColumnMappingId,SourceColumn,TargetColumn")] ColumnMapping columnMapping)
        {
            if (ModelState.IsValid)
            {
                db.Entry(columnMapping).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(columnMapping);
        }

        // GET: ColumnMappings/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ColumnMapping columnMapping = db.ColumnMapping.Find(id);
            if (columnMapping == null)
            {
                return HttpNotFound();
            }
            return View(columnMapping);
        }

        // POST: ColumnMappings/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            ColumnMapping columnMapping = db.ColumnMapping.Find(id);
            db.ColumnMapping.Remove(columnMapping);
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
