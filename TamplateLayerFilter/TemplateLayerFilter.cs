//AUTOCAD
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.LayerManager;
using Microsoft.Win32;
using System.IO;
using AcadLib;
using System;
using Autodesk.AutoCAD.Windows;

namespace Vil.Acad.TemplateLayerFilter
{
    public static class Commands
    {
        [CommandMethod(nameof(Pik_ImportLayerFilters))]
        public static void Pik_ImportLayerFilters()
        {
            CommandStart.Start(doc =>
            {
                // Выбор файла
                var file = SelectFile();
                LayerFilterTree lft = doc.Database.LayerFilters;
                ImportLayerFilterTree(file, lft.Root, doc.Database);
                doc.Database.LayerFilters = lft;
            });
        }        

        [CommandMethod("TemplateLayerFilter")]
        public static void TemplateLayerFilter()
        {
            CommandStart.Start(doc =>
            {
                Database db = doc.Database;
                Editor ed = doc.Editor;

                string qNewTemplateFile = GetQNewTemplateFile();
                if (qNewTemplateFile == string.Empty || !File.Exists(qNewTemplateFile))
                {
                    ed.WriteMessage("\nНе определен шаблон по умолчанию, из которого должны корпироваться фильтры слоев.");
                    return;
                }

                LayerFilterTree lft = db.LayerFilters;
                ImportLayerFilterTree(qNewTemplateFile, lft.Root, db);
                db.LayerFilters = lft;
            });            
        }

        private static string GetQNewTemplateFile()
        {
            string res = string.Empty;
            RegistryKey regKey = Registry.CurrentUser.OpenSubKey(HostApplicationServices.Current.UserRegistryProductRootKey);
            regKey = regKey.OpenSubKey("Profiles");
            regKey = regKey.OpenSubKey(regKey.GetValue(null).ToString());
            regKey = regKey.OpenSubKey("General");
            res = regKey.GetValue("QnewTemplate").ToString();
            return res;
        }

        private static void ImportLayerFilterTree(string filePath, LayerFilter lfDest, Database dbDest)
        {
            using (Database dbSrc = new Database(false, false))
            {
                dbSrc.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, false, string.Empty);
                dbSrc.CloseInput(true);
                ImportNestedFilters(dbSrc.LayerFilters.Root, lfDest, dbSrc, dbDest);
            }
        }

        public static void ImportNestedFilters(LayerFilter srcFilter, LayerFilter destFilter, Database srcDb, Database destDb)
        {
            using (Transaction t = srcDb.TransactionManager.StartTransaction())
            {
                LayerTable lt = t.GetObject(srcDb.LayerTableId, OpenMode.ForRead, false) as LayerTable;

                foreach (LayerFilter sf in srcFilter.NestedFilters)
                {
                    // Получаем слои, которые следует клонировать в db
                    // Только те, которые участвуют в фильтре
                    ObjectIdCollection layerIds = new ObjectIdCollection();
                    foreach (ObjectId layerId in lt)
                    {
                        LayerTableRecord ltr = t.GetObject(layerId, OpenMode.ForRead, false) as LayerTableRecord;
                        if (sf.Filter(ltr))
                        {
                            layerIds.Add(layerId);
                        }
                    }

                    // Клонируем слои во внешнюю db 
                    IdMapping idmap = new IdMapping();
                    if (layerIds.Count > 0)
                    {
                        srcDb.WblockCloneObjects(layerIds, destDb.LayerTableId, idmap, DuplicateRecordCloning.Replace, false);
                    }

                    // Опеределяем не было ли фильтра слоев
                    // с таким же именем во внешней db
                    LayerFilter df = null;
                    foreach (LayerFilter f in destFilter.NestedFilters)
                    {
                        if (f.Name.Equals(sf.Name))
                        {
                            df = f;
                            break;
                        }
                    }

                    if (df == null)
                    {
                        if (sf is LayerGroup)
                        {
                            // Создаем новую группу слоев если
                            // ничего не найдено
                            LayerGroup sfgroup = sf as LayerGroup;
                            LayerGroup dfgroup = new LayerGroup();
                            dfgroup.Name = sf.Name;

                            df = dfgroup;

                            LayerCollection lyrs = sfgroup.LayerIds;
                            foreach (ObjectId lid in lyrs)
                            {
                                if (idmap.Contains(lid))
                                {
                                    IdPair idp = idmap[lid];
                                    dfgroup.LayerIds.Add(idp.Value);
                                }
                            }
                            destFilter.NestedFilters.Add(df);
                        }
                        else
                        {
                            // Создаем фильтр слоев если
                            // ничего не найдено
                            df = new LayerFilter();
                            df.Name = sf.Name;
                            df.FilterExpression = sf.FilterExpression;
                            destFilter.NestedFilters.Add(df);
                        }
                    }

                    // Импортируем другие фильтры
                    ImportNestedFilters(sf, df, srcDb, destDb);
                }
                t.Commit();
            }
        }

        private static string SelectFile()
        {
            var dlg = new OpenFileDialog("Выбор файла с фильтрами слоев", "", "dwg; dwt", "dlgName", OpenFileDialog.OpenFileDialogFlags.NoUrls);
            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                throw new CancelByUserException();
            }
            return dlg.Filename;
        }
    }
}
