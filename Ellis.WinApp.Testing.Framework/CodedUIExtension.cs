//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UITesting;

namespace Ellis.WinApp.Testing.Framework
{
    public static class CodedUIExtension
    {
        public static T SearchFor<T>(this UITestControl _this, dynamic searchProperties, dynamic filterProperties = null)
            where T : UITestControl, new()
        {
            var ctrl = new T {Container = _this};
            var propNames = ((object) searchProperties).GetPropertiesForObject();

            foreach (var item in propNames)
            {
                ctrl.TechnologyName = "MSAA";
                ctrl.SearchProperties.Add(item, ((object) searchProperties).GetPropertyValue(item).ToString());
            }

            if (filterProperties != null)
            {
                propNames = ((object) filterProperties).GetPropertiesForObject();
                foreach (var item in propNames)
                {
                    ctrl.TechnologyName = "MSAA";
                    ctrl.SearchProperties.Add(item, ((object) filterProperties).GetPropertyValue(item).ToString());
                }
            }

            return ctrl;
        }

        //public static List<T> SearchForAll<T>(this UITestControl _this, dynamic searchProperties,
        //    dynamic filterProperties = null)
        //    where T : UITestControl, new()
        //{
        //    var ctrl = new T();
        //    var controls = new List<T>();
        //    ctrl.Container = _this;
        //    var propNames = ((object) searchProperties).GetPropertiesForObject();

        //    foreach (var item in propNames)
        //    {
        //        ctrl.TechnologyName = "MSAA";
        //        ctrl.SearchProperties.Add(item, ((object) searchProperties).GetAllPropertyValues(item).ToString());
        //        controls.Add(ctrl);
        //    }

        //    if (filterProperties != null)
        //    {
        //        propNames = ((object) filterProperties).GetPropertiesForObject();
        //        foreach (var item in propNames)
        //        {
        //            ctrl.TechnologyName = "MSAA";
        //            ctrl.SearchProperties.Add(item, ((object) filterProperties).GetAllPropertyValues(item).ToString());
        //            controls.Add(ctrl);
        //        }
        //    }

        //    return controls;
        //}

        private static IEnumerable<string> GetPropertiesForObject(this object _this)
        {
            return (from x in _this.GetType().GetProperties() select x.Name).ToList();
        }

        private static object GetPropertyValue(this object _this, string propName)
        {
            var prop = (from x in _this.GetType().GetProperties() where x.Name == propName select x).FirstOrDefault();
            return prop.GetValue(_this);
        }

        private static List<object> GetAllPropertyValues(this object _this, string propName)
        {
            var props = (from x in _this.GetType().GetProperties() where x.Name == propName select x).ToList();
            List<object> objProp = null;
            objProp.AddRange(props.Select(prop => prop.GetValue(_this)));
            return objProp;
        }
    }
}