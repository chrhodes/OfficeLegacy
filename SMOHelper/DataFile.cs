using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class DataFile
    {
        #region Initialization

        /// <summary>
        /// Initializes a new instance of the DataFile class.
        /// </summary>
        public DataFile(SMO.DataFile dataFile)
        {
            try { AvailableSpace = dataFile.AvailableSpace.ToString(); }            catch(Exception) { AvailableSpace = "Not Available"; }
            try { FileName = dataFile.FileName.ToString(); }                        catch(Exception) { FileName = "Not Available"; }
            try { Growth = dataFile.Growth.ToString(); }                            catch(Exception) { Growth = "Not Available"; }
            try { GrowthType = dataFile.GrowthType.ToString(); }                    catch(Exception) { GrowthType = "Not Available"; }
            try { ID = dataFile.ID.ToString(); }                                    catch(Exception) { ID = "Not Available"; }
            try { IsOffline = dataFile.IsOffline.ToString(); }                      catch(Exception) { IsOffline = "Not Available"; }
            try { IsPrimaryFile = dataFile.IsPrimaryFile.ToString(); }              catch(Exception) { IsPrimaryFile = "Not Available"; }
            try { IsReadOnly = dataFile.IsReadOnly.ToString(); }                    catch(Exception) { IsReadOnly = "Not Available"; }
            try { IsReadOnlyMedia = dataFile.IsReadOnlyMedia.ToString(); }          catch(Exception) { IsReadOnlyMedia = "Not Available"; }
            try { IsSparse = dataFile.IsSparse.ToString(); }                        catch(Exception) { IsSparse = "Not Available"; }
            //try { IsTouched = dataFile.IsTouched.ToString(); }                    catch(Exception) { IsTouched = "Not Available"; }
            try { MaxSize = dataFile.MaxSize.ToString(); }                          catch(Exception) { MaxSize = "Not Available"; }
            try { Name = dataFile.Name.ToString(); }                                catch(Exception) { Name = "Not Available"; }
            try { NumberOfDiskReads = dataFile.NumberOfDiskReads.ToString(); }      catch(Exception) { NumberOfDiskReads = "Not Available"; }
            try { NumberOfDiskWrites = dataFile.NumberOfDiskWrites.ToString(); }    catch(Exception) { NumberOfDiskWrites = "Not Available"; }
            try { Size = dataFile.Size.ToString(); }                                catch(Exception) { Size = "Not Available"; }
            try { State = dataFile.State.ToString(); }                              catch(Exception) { State = "Not Available"; }
            try { UsedSpace = dataFile.UsedSpace.ToString(); }                      catch(Exception) { UsedSpace = "Not Available"; }
            try { VolumeFreeSpace = dataFile.VolumeFreeSpace.ToString(); }          catch(Exception) { VolumeFreeSpace = "Not Available"; }
        }

        #endregion

        public string AvailableSpace;
        public string FileName;
        public string Growth;
        public string GrowthType;
        public string ID;
        public string IsOffline;
        public string IsPrimaryFile;
        public string IsReadOnly;
        public string IsReadOnlyMedia;
        public string IsSparse;
        public string IsTouched;
        public string MaxSize;
        public string Name;
        public string NumberOfDiskReads;
        public string NumberOfDiskWrites;
        public string Size;
        public string State;
        public string UsedSpace;
        public string VolumeFreeSpace;

    }
}
