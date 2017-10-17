/**************************************
 *
 * Part of OLEWOO - http://www.benf.org (2010-2012)
 *
 * CopyLeft, but please credit.
 *
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using olewoo_interop;

namespace olewoo_cs
{
    class IDLGrabber : olewoo_interop.IDLFormatter_iop
    {
        String _s;
        public override void AddLink(string s, string s2)
        {
            _s += s;
        }
        public override void AddString(string s)
        {
            _s += s;
        }
        public String Value
        {
            get
            {
                return _s;
            }
        }
    }

    static class ITypeInfoXtra
    {

        public static String PaddedHex(this int x)
        {
            return "0x" + x.ToString("x").PadLeft(8, '0');
        }
        const int MEMBERID_NONE = -1;
        public static String GetName(this ITypeLib ti)
        {
            String res;
            String ignored;
            String ignored2;
            int ctx;
            ti.GetDocumentation(MEMBERID_NONE, out res, out ignored, out ctx, out ignored2);
            return res;
        }
        public static String GetName(this ITypeInfo ti)
        {
            return ti.GetDocumentationById(MEMBERID_NONE);
        }
        public static String GetDocumentationById(this ITypeInfo ti, int memid)
        {
            String res;
            String ignored;
            String ignored2;
            int ctx;
            ti.GetDocumentation(memid, out res, out ignored, out ctx, out ignored2);
            return res;
        }
        public static String GetHelpDocumentationById(this ITypeInfo ti, int memid, out int context)
        {
            String res;
            String ignored;
            String ignored2;
            ti.GetDocumentation(memid, out ignored, out res, out context, out ignored2);
            return res;
        }
        public static String GetHelpDocumentation(this ITypeLib ti, out int context)
        {
            String res;
            String ignored;
            String ignored2;
            ti.GetDocumentation(MEMBERID_NONE, out ignored, out res, out context, out ignored2);
            return res;
        }
        public static TypeAttr GetTypeAttr(this ITypeInfo ti)
        {         
            TypeAttr ta = new TypeAttr(ti);
            return ta;
        }
        public static bool SwapForInterface(ref ITypeInfo ti, ref TypeAttr ta)
        {
            try
            {
                if (ta.typekind == TypeAttr.TypeKind.TKIND_DISPATCH && 0 != (ta.wTypeFlags & TypeAttr.TypeFlags.TYPEFLAG_FDUAL))
                {
                    int href;
                    ti.GetRefTypeOfImplType(-1, out href);
                    ti.GetRefTypeInfo(href, out ti);
                    ta = new TypeAttr(ti);
                    return true;
                }
            }
            catch (Exception)
            {
            }
            return false;
        }

        public static String QuoteString(Object o)
        {
            if (o == null) return "";
            if (o.GetType() == typeof(System.String))
            {
                return "\"" + o.ToString() + "\"";
            }
            return o.ToString();
        }

        public static String ReEscape(this String s)
        {
            if (s.IndexOfAny(new char[]{'\0', '\n', '\r', '\b', '\a', '\f', '\t', '\v'}) >= 0)
            {
                s = s.Replace("\0", "\\0");
                s = s.Replace("\n", "\\n");
                s = s.Replace("\r", "\\r");
                s = s.Replace("\b", "\\b");
                s = s.Replace("\a", "\\a");
                s = s.Replace("\f", "\\f");
                s = s.Replace("\t", "\\t");
                s = s.Replace("\v", "\\v");
            }
            return "\"" + s + "\"";
        }

        public static String VarTypeToString(long vt)
        {
            switch ((VarEnum)vt)
            {
                case VarEnum.VT_EMPTY: // 0,
                case VarEnum.VT_NULL: // 1,
                case VarEnum.VT_I2: // 2,
                case VarEnum.VT_I4: // 3,
                case VarEnum.VT_R4: // 4,
                case VarEnum.VT_R8: // 5,
                case VarEnum.VT_CY: // 6,
                case VarEnum.VT_DATE: // 7,
                case VarEnum.VT_BSTR: // 8,
                case VarEnum.VT_DISPATCH: // 9,
                case VarEnum.VT_ERROR: // 10,
                case VarEnum.VT_BOOL: // 11,
                case VarEnum.VT_VARIANT: // 12,
                case VarEnum.VT_UNKNOWN: // 13,
                case VarEnum.VT_DECIMAL: // 14,
                case VarEnum.VT_I1: // 16,
                case VarEnum.VT_UI1: // 17,
                case VarEnum.VT_UI2: // 18,
                    break;
                case VarEnum.VT_UI4: // 19,
                    return "unsigned long";
                case VarEnum.VT_I8: // 20,
                    return "int64";
                case VarEnum.VT_UI8: // 21,
                    return "uint64";
                case VarEnum.VT_INT: // 22,
                case VarEnum.VT_UINT: // 23,
                case VarEnum.VT_VOID: // 24,
                case VarEnum.VT_HRESULT: // 25,
                case VarEnum.VT_PTR: // 26,
                case VarEnum.VT_SAFEARRAY: // 27,
                case VarEnum.VT_CARRAY: // 28,
                    break;
                case VarEnum.VT_USERDEFINED: // 29,
                    return "USER DEFINED";
                case VarEnum.VT_LPSTR: // 30,
                case VarEnum.VT_LPWSTR: // 31,
                    return "LPWSTR";
                case VarEnum.VT_RECORD: // 36,
                //            case VarEnum.VT_INT_PTR	: // 37,
                //            case VarEnum.VT_UINT_PTR	: // 38,
                case VarEnum.VT_FILETIME: // 64,
                case VarEnum.VT_BLOB: // 65,
                case VarEnum.VT_STREAM: // 66,
                case VarEnum.VT_STORAGE: // 67,
                case VarEnum.VT_STREAMED_OBJECT: // 68,
                case VarEnum.VT_STORED_OBJECT: // 69,
                case VarEnum.VT_BLOB_OBJECT: // 70,
                case VarEnum.VT_CF: // 71,
                case VarEnum.VT_CLSID: // 72,
                //            case VarEnum.VT_VERSIONED_STREAM	: // 73,
                //            case VarEnum.VT_BSTR_BLOB	: // 0xfff,
                case VarEnum.VT_VECTOR: // 0x1000,
                    break;
            }
            return vt.ToString() + "???";
        }

        public static String[] GetNames(this FuncDesc fd, ITypeInfo ti)
        {
            String[] names = new String[fd.cParams + 1];
            int pcNames;
            ti.GetNames(fd.memid, names, fd.cParams + 1, out pcNames);
            return names;
        }
    }

    public abstract class ITlibNode
    {
        public delegate List<ITlibNode> dlgCreateChildren();

        public enum ImageIndices
        {
            idx_coclass,
            idx_const,
            idx_dispinterface,
            idx_enum,
            idx_interface,
            idx_strucmem,
            idx_struct,
            idx_typelib,
            idx_methodlist,
            idx_method,
            idx_typedef,
            idx_module,
            idx_constlist,
            idx_selected,
            idx_propertylist
        };

        private List<ITlibNode> _children; 

        public abstract ITlibNode Parent
        {
            get;
        }
        public List<ITlibNode> Children
        {
            get
            {
                if (_children == null)
                {
                    _children = GenChildren();
                    for (int i = 0; i < _children.Count; ++i) _children[i].Idx = i;
                }
                return _children;
            }
            set { _children = value; }
        }

        public abstract List<ITlibNode> GenChildren();
        public abstract String Name
        {
            get;
        }
        public int Idx
        {
            get;
            set;
        }
        public abstract String ShortName
        {
            get;
        }
        public abstract String ObjectName
        {
            get;
        }
        public abstract bool DisplayAtTLBLevel(ICollection<String> interfaceNames);
        public abstract int ImageIndex { get; }
        public void CommonBuildTlibNode(ITlibNode parent, ITypeInfo ti, bool topLevel, bool swapfordispatch, List<ITlibNode> res)
        {
            TypeAttr ta = ti.GetTypeAttr();
            switch (ta.typekind)
            {
                case TypeAttr.TypeKind.TKIND_DISPATCH:
                    res.Add(new OWDispInterface(this, ti, ta, topLevel));
                    if (swapfordispatch && ITypeInfoXtra.SwapForInterface(ref ti, ref ta))
                    {
                        res.Add(new OWInterface(this, ti, ta, topLevel));
                    }
                    break;
                case TypeAttr.TypeKind.TKIND_INTERFACE:
                    res.Add(new OWInterface(this, ti, ta, topLevel));
                    break;
                case TypeAttr.TypeKind.TKIND_ALIAS:
                    res.Add(new OWTypeDef(this, ti, ta));
                    break;
                case TypeAttr.TypeKind.TKIND_ENUM:
                    res.Add(new OWEnum(this, ti, ta));
                    break;
                case TypeAttr.TypeKind.TKIND_COCLASS:
                    res.Add(new OWCoClass(this, ti, ta));
                    break;
                case TypeAttr.TypeKind.TKIND_RECORD:
                    res.Add(new OWRecord(this, ti, ta));
                    break;
                case TypeAttr.TypeKind.TKIND_MODULE:
                    res.Add(new OWModule(this, ti, ta));
                    break;
                default:
                    break;
            }
        }

        public abstract void BuildIDLInto(IDLFormatter ih);

        protected void AddHelpStringAndContext(List<String> lprops, String help, int context)
        {
            if (!String.IsNullOrEmpty(help)) lprops.Add("helpstring(\"" + help + "\")");
            if (context != 0) lprops.Add("helpcontext(" + context.PaddedHex() + ")");

        }
    }

    public interface IClearUp
    {
        void ClearUp();
    }

    public class OWTypeLib : ITlibNode, IClearUp
    {
        ITypeLib _tlib;
        String _name;

        public OWTypeLib(String path)
        {
            NativeMethods.LoadTypeLib(path, out _tlib);
            if (_tlib == null) throw new Exception(path + " is not a loadable typelibrary.");
            _name = _tlib.GetName();
            int cnt;
            _name += " (" + _tlib.GetHelpDocumentation(out cnt) + ")";
        }

        public void ClearUp()
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_tlib);
            _tlib = null; 
        }

        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            throw new Exception("Should not be calling this!"); 
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_typelib; }
        }
        public override String ShortName
        {
            get
            {
                return _tlib.GetName();
            }
        }
        public override string ObjectName
        {
            get { return ShortName; }
        }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            int ticount = _tlib.GetTypeInfoCount();
            for (int x = 0; x < ticount; ++x)
            {
                ITypeInfo ti;
                _tlib.GetTypeInfo(x, out ti);
                CommonBuildTlibNode(this, ti, true, true, res);
            }
            return res;
        }

        public override String Name
        {
            get
            {
                return _name;
            }
        }

        public override ITlibNode Parent
        {
            get { return null; }
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            // Header for type library, followed by a first pass to pre-declare
            // interfaces.
            // dispinterfaces aren't shown seperately.

            ih.AppendLine("// Generated .IDL file (by OleWoo)");
            ih.AppendLine("[");
            List<String> liba = new List<string>();
            using (TypeLibAttr tla = new TypeLibAttr(_tlib))
            {
                liba.Add("uuid(" + tla.guid + ")");
                liba.Add("version(" + tla.wMajorVerNum + "." + tla.wMinorVerNum + ")");
            }
            CustomDatas cds = new CustomDatas(_tlib as ITypeLib2);
            {
                foreach (CustomData cd in cds.Items)
                {
                    liba.Add("custom(" + cd.guid + ", " + ITypeInfoXtra.QuoteString(cd.varValue) + ")");
                }
            }
            int cnt = 0;
            String help = _tlib.GetHelpDocumentation(out cnt);
            if (!String.IsNullOrEmpty(help)) liba.Add("helpstring(\"" + help + "\")");
            if (cnt != 0) liba.Add("helpcontext(" + cnt.PaddedHex() + ")");

            cnt = 0;
            liba.ForEach( x => ih.AppendLine("  " + x + (++cnt == liba.Count ? "" : ",")) );
            ih.AppendLine("]");
            ih.AppendLine("library " + ShortName);
            ih.AppendLine("{");
            using (new IDLHelperTab(ih))
            {
                // How do I know I'm importing stdole2??!
                // Forward declare all interfaces.
                ih.AppendLine("// Forward declare all types defined in this typelib");
                /* 
                 * Need to collect all dumpable interface names, in case we have dispinterfaces which don't have
                 * top level interfaces.  In THIS case, we'd dump the dispinterface.
                 */
                ICollection<String> interfaceNames = Children.Aggregate<ITlibNode, ICollection<String>>(new HashSet<String>(),
                    (x, y) =>
                    {
                        if ((y as OWInterface) != null) x.Add(y.ShortName);
                        return x;
                    });
                Children.FindAll(x => ((x as OWInterface) != null || (x as OWDispInterface) != null)).ForEach(
                    x => ih.AppendLine(x.Name)
                        );
                Children.FindAll(x => x.DisplayAtTLBLevel(interfaceNames)).ForEach(
                   x => { x.BuildIDLInto(ih); ih.AppendLine(""); }
                );
            }
            ih.AppendLine("};");
        }

    }

    class OWCoClass : ITlibNode
    {
        ITlibNode _parent;
        String _name;
        TypeAttr _ta;
        ITypeInfo _ti;
        public OWCoClass(ITlibNode parent, ITypeInfo ti, TypeAttr ta)
        {
            _parent = parent;
            _name = ti.GetName();
            _ta = ta;
            _ti = ti;
        }
        public override String Name
        {
            get
            {
                return "coclass " + _name;
            }
        }
        public override string ObjectName
        {
            get { return _name + "#c"; }
        }
        public override string ShortName
        {
            get { return _name; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return true;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_coclass; }
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            for (int x = 0; x < _ta.cImplTypes; ++x)
            {
                int href;
                ITypeInfo ti2;
                _ti.GetRefTypeOfImplType(x, out href);
                _ti.GetRefTypeInfo(href, out ti2);
                CommonBuildTlibNode(this, ti2, false, false, res);
            }
            return res;
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            ih.AppendLine("[");
            List<String> lprops = new List<string>();
            lprops.Add("uuid(" + _ta.guid + ")");
            int context;
            String help = _ti.GetHelpDocumentationById(-1, out context);
            AddHelpStringAndContext(lprops, help, context);
            for (int i = 0; i < lprops.Count; ++i)
            {
                ih.AppendLine("  " + lprops[i] + (i < (lprops.Count - 1) ? "," : ""));
            }
            ih.AppendLine("]");
            ih.AppendLine("coclass " + _name + " {");
            using (new IDLHelperTab(ih))
            {
                for (int x = 0; x < _ta.cImplTypes; ++x)
                {
                    int href;
                    ITypeInfo ti2;
                    _ti.GetRefTypeOfImplType(x, out href);
                    _ti.GetRefTypeInfo(href, out ti2);
                    System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS itypflags;
                    _ti.GetImplTypeFlags(x, out itypflags);
                    List<String> res = new List<String>();
                    if (0 != (itypflags & System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT)) res.Add("default");
                    if (0 != (itypflags & System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE)) res.Add("source");

                    if (res.Count > 0) ih.AddString("[" + String.Join(", ", res.ToArray()) + "] ");
                    ih.AddString("interface ");
                    ih.AddLink(ti2.GetName(), "i");
                    ih.AppendLine(";");
                }
            }
            ih.AppendLine("};");
        }
    }

    class OWInterface : ITlibNode
    {
        ITlibNode _parent;
        String _name;
        TypeAttr _ta;
        ITypeInfo _ti;
        bool _topLevel;

        public OWInterface(ITlibNode parent, ITypeInfo ti, TypeAttr ta, bool topLevel)
        {
            _parent = parent;
            _name = ti.GetName();
            _ta = ta;
            _ti = ti;
            _topLevel = topLevel;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_interface; }
        }
        public override String Name
        {
            get
            {
                return (_topLevel ? "interface " : "") + _name;
            }
        }
        public override string ObjectName
        {
            get { return _name + "#i"; }
        }
        public override string ShortName
        {
            get { return _name; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return true;
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            // First add a child for every method / property, then an inherited interfaces
            // child (if applicable).

            int nfuncs = _ta.cFuncs;
            for (int idx = 0; idx < nfuncs; ++idx)
            {
                FuncDesc fd = new FuncDesc(_ti, idx);
                res.Add(new OWMethod(this, _ti, fd));
            }
            if (_ta.cImplTypes > 0)
            {
                res.Add(new OWInheritedInterfaces(this, _ti, _ta));
            }
            return res;
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            ih.AppendLine("[");
            List<String> lprops = new List<string>();
            lprops.Add("uuid(" + _ta.guid + ")");
            int context;
            String help = _ti.GetHelpDocumentationById(-1, out context);
            AddHelpStringAndContext(lprops, help, context);
            if (0 != (_ta.wTypeFlags & TypeAttr.TypeFlags.TYPEFLAG_FHIDDEN)) lprops.Add("hidden");
            if (0 != (_ta.wTypeFlags & TypeAttr.TypeFlags.TYPEFLAG_FDUAL)) lprops.Add("dual");
            if (0 != (_ta.wTypeFlags & TypeAttr.TypeFlags.TYPEFLAG_FRESTRICTED)) lprops.Add("restricted");
            if (0 != (_ta.wTypeFlags & TypeAttr.TypeFlags.TYPEFLAG_FNONEXTENSIBLE)) lprops.Add("nonextensible");
            if (0 != (_ta.wTypeFlags & TypeAttr.TypeFlags.TYPEFLAG_FOLEAUTOMATION)) lprops.Add("oleautomation");
            for (int i = 0; i < lprops.Count; ++i)
            {
                ih.AppendLine("  " + lprops[i] + (i < (lprops.Count - 1) ? "," : ""));
            }
            ih.AppendLine("]");

            if (_ta.cImplTypes > 0)
            {
                int href;
                _ti.GetRefTypeOfImplType(0, out href);
                ITypeInfo ti2;
                _ti.GetRefTypeInfo(href, out ti2);
                ih.AddString("interface " + _name + " : ");
                ih.AddLink(ti2.GetName(), "i");
                ih.AppendLine(" {");
            }
            else
            {
                ih.AppendLine("interface " + _name + " {");
            }
            using (new IDLHelperTab(ih))
            {
                Children.ForEach( x => x.BuildIDLInto(ih) );
            }
            ih.AppendLine("};");
        }
    }

    class OWInheritedInterfaces : ITlibNode
    {
        ITlibNode _parent;
        protected TypeAttr _ta;
        protected ITypeInfo _ti;
        public OWInheritedInterfaces(ITlibNode parent, ITypeInfo ti, TypeAttr ta)
        {
            _parent = parent;
            _ta = ta;
            _ti = ti;
        }
        public override String Name
        {
            get
            {
                return "Inherited Interfaces";
            }
        }
        public override string ShortName
        {
            get { return null; }
        }
        public override string ObjectName
        {
            get { return null; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return false;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_interface; }
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();

            if (_ta.cImplTypes > 0)
            {
                if (_ta.cImplTypes > 1) throw new Exception("Multiple inheritance!?");
                int href;
                ITypeInfo ti2;
                _ti.GetRefTypeOfImplType(0, out href);
                _ti.GetRefTypeInfo(href, out ti2);
                CommonBuildTlibNode(this, ti2, false, true, res);
            }
            return res;
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            ih.AppendLine("");
        }
    }

    /*
     * A dispinterface's first inherited interface is the swap for interface.
     */
    class OWDispInterfaceInheritedInterfaces : OWInheritedInterfaces
    {
        public OWDispInterfaceInheritedInterfaces(ITlibNode parent, ITypeInfo ti, TypeAttr ta) 
            : base(parent, ti, ta)
        {
        }
        public override List<ITlibNode> GenChildren()
        {
            ITypeInfo ti = _ti;
            TypeAttr ta = _ta;
            ITypeInfoXtra.SwapForInterface(ref ti, ref ta);
            List<ITlibNode> res = new List<ITlibNode>();
            res.Add(new OWInterface(this, ti, ta, false));
            return res;
        }
    }

    class OWMethod : ITlibNode
    {
        ITlibNode _parent;
        String _name;
        FuncDesc _fd;
        ITypeInfo _ti;

        public OWMethod(ITlibNode parent, ITypeInfo ti, FuncDesc fd)
        {
            _parent = parent;
            _ti = ti;
            _fd = fd;

            String[] names = fd.GetNames(ti);
            String functionname = names[0];

            _name = names[0];
        }
        public override String Name
        {
            get
            {
                return _name;
            }
        }
        public override string ShortName
        {
            get { return _name; }
        }
        public override string ObjectName
        {
            get { return _name + "#m"; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return false;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_method; }
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            return res;
        }
        private String ParamFlagsDescription(ParamDesc pd)
        {
            ParamDesc.ParamFlags flg = pd.wParamFlags;
            List<String> res = new List<string>();
            if (0 != (flg & ParamDesc.ParamFlags.PARAMFLG_FIN)) res.Add("in");
            if (0 != (flg & ParamDesc.ParamFlags.PARAMFLG_FOUT)) res.Add("out");
            if (0 != (flg & ParamDesc.ParamFlags.PARAMFLG_FRETVAL)) res.Add("retval");
            if (0 != (flg & ParamDesc.ParamFlags.PARAMFLG_FOPT)) res.Add("optional");
            if (0 != (flg & ParamDesc.ParamFlags.PARAMFLG_FHASDEFAULT)) res.Add("defaultvalue(" + ITypeInfoXtra.QuoteString(pd.varDefaultValue) + ")");
            return "[" + String.Join(", ", res.ToArray()) + "]";
        }
        public bool Property 
        {
            get
            {
                return _fd.invkind != FuncDesc.InvokeKind.INVOKE_FUNC;
            }
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            BuildIDLInto(ih, false);
        }
        delegate void GenPt(int x);
        public void BuildIDLInto(IDLFormatter ih, bool bAsDispatch)
        {
            bool memIdInSpecialRange = (_fd.memid >= 0x60000000 && _fd.memid < 0x60020000);
            List<String> lprops = new List<string>();
            if (!memIdInSpecialRange)
            {
                lprops.Add("id(" + _fd.memid.PaddedHex() + ")");
            }
            switch (_fd.invkind)
            {
                case FuncDesc.InvokeKind.INVOKE_PROPERTYGET:
                    lprops.Add("propget");
                    break;
                case FuncDesc.InvokeKind.INVOKE_PROPERTYPUT:
                    lprops.Add("propput");
                    break;
                case FuncDesc.InvokeKind.INVOKE_PROPERTYPUTREF:
                    lprops.Add("propputref");
                    break;
            }
            int context;
            String help = _ti.GetHelpDocumentationById(_fd.memid, out context);
            if (0 != (_fd.wFuncFlags & FuncDesc.FuncFlags.FUNCFLAG_FRESTRICTED)) lprops.Add("restricted");
            if (0 != (_fd.wFuncFlags & FuncDesc.FuncFlags.FUNCFLAG_FHIDDEN)) lprops.Add("hidden");
            AddHelpStringAndContext(lprops, help, context);
            ih.AppendLine("[" + String.Join(", ", lprops.ToArray()) + "] ");
            // Prototype in a different line.
            ElemDesc ed = _fd.elemdescFunc;

            GenPt paramtextgen = null;
            ElemDesc elast = null;
            bool bRetvalPresent = false;
            if (_fd.cParams > 0)
            {
                String[] names = _fd.GetNames(_ti);
                ElemDesc[] edps = _fd.elemdescParams;
                if (edps.Length > 0) elast = edps[edps.Length - 1];
                if (bAsDispatch && elast != null && 0 != (elast.paramdesc.wParamFlags & ParamDesc.ParamFlags.PARAMFLG_FRETVAL))
                    bRetvalPresent = true;
                int maxCnt = (bAsDispatch && bRetvalPresent) ? _fd.cParams - 1 : _fd.cParams;

                paramtextgen = x =>
                    {
                        String paramname = (names[x + 1] == null) ? "rhs" : names[x + 1];
                        ElemDesc edp = edps[x];
                        ParamDesc fd = edp.paramdesc;
                        ih.AddString(ParamFlagsDescription(edp.paramdesc) + " ");
                        edp.tdesc.ComTypeNameAsString(_ti, ih);
                        ih.AddString(" " + paramname);
                    }
                    ;
            }
            (bRetvalPresent ? elast : ed).tdesc.ComTypeNameAsString(_ti, ih);
            if (memIdInSpecialRange)
            {
                ih.AddString(" " + _fd.callconv.ToString().Substring(2).ToLower());
            }
            ih.AddString(" " + _name);
            switch (_fd.cParams)
            {
                case 0:
                    ih.AppendLine("();");
                    break;
                case 1:
                    ih.AddString("(");
                    paramtextgen(0);
                    ih.AppendLine(");");
                    break;
                default:
                    ih.AppendLine("(");
                    using (new IDLHelperTab(ih))
                    {
                        for (int y = 0; y < _fd.cParams; ++y)
                        {
                            paramtextgen(y);
                            ih.AppendLine(y == _fd.cParams - 1 ? "" : ",");
                        }
                    }
                    ih.AppendLine(");");
                    break;
            }
        }
    }


    class OWDispInterface : ITlibNode
    {
        ITlibNode _parent;
        String _name;
        TypeAttr _ta;
        ITypeInfo _ti;
        bool _topLevel;

        OWIDispatchMethods _methodChildren;
        OWIDispatchProperties _propChildren;

        public OWDispInterface(ITlibNode parent, ITypeInfo ti, TypeAttr ta, bool topLevel)
        {
            _parent = parent;
            _name = ti.GetName();
            _ta = ta;
            _ti = ti;
            _topLevel = topLevel;
        }
        public override String Name
        {
            get
            {
                return (_topLevel ? "dispinterface " : "") + _name;
            }
        }
        public override string ShortName
        {
            get { return _name; }
        }
        public override string ObjectName
        {
            get { return _name + "#di"; }
        }
        /* Don't show a dispinterface at top level, UNLESS the corresponding interface is not itself
         * at top level. 
         */
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return !(interfaceNames.Contains(ShortName));
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_dispinterface; }
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            if (_ta.cVars > 0) {
                _propChildren = new OWIDispatchProperties(this);
                res.Add(_propChildren);
            }
            if (_ta.cFuncs > 0)
            {
                _methodChildren = new OWIDispatchMethods(this);
                res.Add(_methodChildren);
            }
            if (_ta.cImplTypes > 0)
            {
                res.Add(new OWDispInterfaceInheritedInterfaces(this, _ti, _ta));
            }
            return res;
        }
        public List<ITlibNode> MethodChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            int nfuncs = _ta.cFuncs;
            for (int idx = 0; idx < nfuncs; ++idx)
            {
                FuncDesc fd = new FuncDesc(_ti, idx);
//                if (0 == (fd.wFuncFlags & FuncDesc.FuncFlags.FUNCFLAG_FRESTRICTED))
                    res.Add(new OWMethod(this, _ti, fd));
            }
            return res;
        }
        public List<ITlibNode> PropertyChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            for (int x = 0; x < _ta.cVars; ++x)
            {
                VarDesc vd = new VarDesc(_ti, x);
                res.Add(new OWDispProperty(this, _ti, vd));
            }
            return res;
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            ih.AppendLine("[");
            List<String> lprops = new List<string>();
            lprops.Add("uuid(" + _ta.guid + ")");
            int context;
            String help = _ti.GetHelpDocumentationById(-1, out context);
            AddHelpStringAndContext(lprops, help, context);
            if (0 != (_ta.wTypeFlags & TypeAttr.TypeFlags.TYPEFLAG_FHIDDEN)) lprops.Add("hidden");
            if (0 != (_ta.wTypeFlags & TypeAttr.TypeFlags.TYPEFLAG_FDUAL)) lprops.Add("dual");
            if (0 != (_ta.wTypeFlags & TypeAttr.TypeFlags.TYPEFLAG_FRESTRICTED)) lprops.Add("restricted");
            if (0 != (_ta.wTypeFlags & TypeAttr.TypeFlags.TYPEFLAG_FNONEXTENSIBLE)) lprops.Add("nonextensible");
            if (0 != (_ta.wTypeFlags & TypeAttr.TypeFlags.TYPEFLAG_FOLEAUTOMATION)) lprops.Add("oleautomation");
            for (int i = 0; i < lprops.Count; ++i)
            {
                ih.AppendLine("  " + lprops[i] + (i < (lprops.Count - 1) ? "," : ""));
            }
            ih.AppendLine("]");

            ih.AppendLine("dispinterface " + _name + " {");

            if (_ta.cFuncs > 0 || _ta.cVars > 0)
            {
                // Naughty, but rely on side effect of verifying children.
                List<ITlibNode> children = Children;
                using (new IDLHelperTab(ih))
                {
                    if (_propChildren != null)
                    {
                        _propChildren.BuildIDLInto(ih);
                    }
                    if (_methodChildren != null)
                    {
                        _methodChildren.BuildIDLInto(ih);
                    }
                }
            }
            ih.AppendLine("};");
        }
    }

    class OWIDispatchMethods : ITlibNode
    {
        OWDispInterface _parent;

        public OWIDispatchMethods(OWDispInterface parent)
        {
            _parent = parent;
        }
        public override String Name
        {
            get
            {
                return "Methods";
            }
        }
        public override string ShortName
        {
            get { return null; }
        }
        public override string ObjectName
        {
            get { return null; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return false;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_methodlist; }
        }
        public override ITlibNode Parent { get { return _parent as ITlibNode; } }
        public override List<ITlibNode> GenChildren()
        {
            return _parent.MethodChildren();
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            List<OWMethod> meths = Children.ToList().ConvertAll(x => x as OWMethod);
            ih.AppendLine("methods:");
            if (meths.Count > 0) using (new IDLHelperTab(ih)) meths.ForEach(x => x.BuildIDLInto(ih, true));
        }
    }

    class OWIDispatchProperties : ITlibNode
    {
        OWDispInterface _parent;

        public OWIDispatchProperties(OWDispInterface parent)
        {
            _parent = parent;
        }
        public override String Name
        {
            get
            {
                return "Properties";
            }
        }
        public override string ShortName
        {
            get { return null; }
        }
        public override string ObjectName
        {
            get { return null; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return false;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_propertylist; }
        }
        public override ITlibNode Parent { get { return _parent as ITlibNode; } }
        public override List<ITlibNode> GenChildren()
        {
            return _parent.PropertyChildren();
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            List<OWDispProperty> props = Children.ToList().ConvertAll(x => x as OWDispProperty);
            ih.AppendLine("properties:");
            if (props.Count > 0) using (new IDLHelperTab(ih)) props.ForEach(x => x.BuildIDLInto(ih, true));
        }
    }

    class OWEnum : ITlibNode
    {
        ITlibNode _parent;
        String _name;
        TypeAttr _ta;
        ITypeInfo _ti;
        public OWEnum(ITlibNode parent, ITypeInfo ti, TypeAttr ta)
        {
            _parent = parent;
            _name = ti.GetName();
            _ta = ta;
            _ti = ti;
        }
        public override String Name
        {
            get
            {
                return "typedef enum " + _name;
            }
        }
        public override string ShortName
        {
            get { return _name; }
        }
        public override string ObjectName
        {
            get { return _name + "#i"; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return false;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_enum; }
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            for (int x = 0; x < _ta.cVars; ++x)
            {
                VarDesc vd = new VarDesc(_ti, x);
                res.Add(new OWEnumValue(this, _ti, vd));
            }
            return res;
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            String tde = "typedef ";
            // If the enum has a uuid, or a version associate with it, we provide that information on the same line.

            if (!_ta.guid.Equals(Guid.Empty))
            {
                tde += "[uuid(" + _ta.guid + "), version(" + _ta.wMajorVerNum + "." + _ta.wMinorVerNum + ")]";
                ih.AppendLine(tde);
                tde = "";
            }
            ih.AppendLine(tde + "enum {");
            using (new IDLHelperTab(ih))
            {
                int idx = 0;
                Children.ForEach(x => (x as OWEnumValue).BuildIDLInto(ih, true, ++idx == _ta.cVars));
            }
            ih.AppendLine("} " + _name + ";");
        }
    }

    class OWEnumValue : ITlibNode
    {
        ITlibNode _parent;
        String _name;
        VarDesc _vd;
        ITypeInfo _ti;
        int _val;
        public OWEnumValue(ITlibNode parent, ITypeInfo ti, VarDesc vd)
        {
            _parent = parent;
            _name = ti.GetDocumentationById(vd.memid);
            _val = (int)vd.varValue;
            _vd = vd;
            _ti = ti;
        }
        public override String Name
        {
            get
            {
                // fixme - look at varkind.
                return "const int " + _name + " = " + _val;
            }
        }
        public override string ShortName
        {
            get { return _name;  }
        }
        public override string ObjectName
        {
            get { return null; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return false;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_const; }
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            return res;
        }
        String negStr(int x)
        {
            return (x < 0) ? ("0x" + x.ToString("X")) : x.ToString();
        }
        public void BuildIDLInto(IDLFormatter ih, bool embedded, bool islast)
        {
            ih.AppendLine("const int " + _ti.GetDocumentationById(_vd.memid) + " = " + negStr(_val) + (embedded ? (islast ? "" : ",") : ";"));
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            BuildIDLInto(ih, false, false);
        }
    }


    class OWDispProperty : ITlibNode
    {
        ITlibNode _parent;
        String _name;
        VarDesc _vd;
        ITypeInfo _ti;
        public OWDispProperty(ITlibNode parent, ITypeInfo ti, VarDesc vd)
        {
            _parent = parent;
            _name = ti.GetDocumentationById(vd.memid);
            _vd = vd;
            _ti = ti;
        }
        public override String Name
        {
            get
            {
                // fixme - look at varkind.
                return _name;
            }
        }
        public override string ShortName
        {
            get { return _name; }
        }
        public override string ObjectName
        {
            get { return null; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return false;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_strucmem; }
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            return res;
        }
        String negStr(int x)
        {
            return (x < 0) ? ("0x" + x.ToString("X")) : x.ToString();
        }
        public void BuildIDLInto(IDLFormatter ih, bool embedded)
        {
            bool memIdInSpecialRange = (_vd.memid >= 0x60000000 && _vd.memid < 0x60020000);
            List<String> lprops = new List<string>();
            if (!memIdInSpecialRange)
            {
                lprops.Add("id(" + _vd.memid.PaddedHex() + ")");
            }
            int context;
            String help = _ti.GetHelpDocumentationById(_vd.memid, out context);
//            if (0 != (_vd.wFuncFlags & FuncDesc.FuncFlags.FUNCFLAG_FRESTRICTED)) lprops.Add("restricted");
//            if (0 != (_vd.wFuncFlags & FuncDesc.FuncFlags.FUNCFLAG_FHIDDEN)) lprops.Add("hidden");
            AddHelpStringAndContext(lprops, help, context);
            ih.AppendLine("[" + String.Join(", ", lprops.ToArray()) + "] ");
            // Prototype in a different line.
            ElemDesc ed = _vd.elemDescVar;

            ed.tdesc.ComTypeNameAsString(_ti, ih);
//            if (memIdInSpecialRange)
//            {
//                ih.AddString(" " + _fd.callconv.ToString().Substring(2).ToLower());
//            }
            ih.AppendLine(" " + _name + ";");
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            BuildIDLInto(ih, false);
        }
    }


    class OWModule : ITlibNode
    {
        ITlibNode _parent;
        String _name;
        ITypeInfo _ti;
        TypeAttr _ta;
        String _dllname;
        public OWModule(ITlibNode parent, ITypeInfo ti, TypeAttr ta)
        {
            _parent = parent;
            _ti = ti;
            _ta = ta;
            _name = _ti.GetName();
            if (_ta.cVars > 0 || _ta.cFuncs > 0)
            {
                olewoo_interop.ITypeInfoXtra tix = new olewoo_interop.ITypeInfoXtra();
                int memid;
                FuncDesc.InvokeKind invkind = FuncDesc.InvokeKind.INVOKE_FUNC;
                if (_ta.cFuncs > 0)
                {
                    FuncDesc fd = new FuncDesc(_ti, 0);
                    invkind = fd.invkind;
                    memid = fd.memid;
                    _dllname = tix.GetDllEntry(ti, invkind, memid);
                }
                else
                {
                    _dllname = null;
                }
            }
        }
        public override String Name
        {
            get
            {
                return "module " + _name;
            }
        }
        public override string ShortName
        {
            get { return _name; }
        }
        public override string ObjectName
        {
            get { return null; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return true;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_module; }
        }
        public override ITlibNode Parent { get { return _parent; } }

        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            if (_ta.cVars > 0) res.Add(new OWChildrenIndirect(this, "Constants", (int)ImageIndices.idx_constlist, GenConstChildren));
            if (_ta.cFuncs > 0) res.Add(new OWChildrenIndirect(this, "Functions", (int)ImageIndices.idx_methodlist, GenFuncChildren));
            return res;
        }
        private List<ITlibNode> GenConstChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            for (int x = 0; x < _ta.cVars; ++x)
            {
                VarDesc vd = new VarDesc(_ti, x);
                res.Add(new OWModuleConst(this, _ti, vd, x));
            }
            return res;
        }
        private List<ITlibNode> GenFuncChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            for (int x = 0; x < _ta.cFuncs; ++x)
            {
                FuncDesc fd = new FuncDesc(_ti, x);
                res.Add(new OWMethod(this, _ti, fd));
            }
            return res;
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            if (_ta.cFuncs == 0)
            {
                ih.AppendLine("// NOTE: This module has no entry points. There is no way to");
                ih.AppendLine("//       extract the dllname of a module with no entry points!");
                ih.AppendLine("// ");
            }
            ih.AppendLine("[");
            List<String> liba = new List<string>();
            liba.Add("dllname(\"" + ( String.IsNullOrEmpty(_dllname) ? "<no entry points>" : _dllname) + "\")");

            if (_ta.guid != Guid.Empty) liba.Add("uuid(" + _ta.guid + ")");
            int cnt = 0;
            String help = _ti.GetHelpDocumentationById(-1, out cnt);
            if (!String.IsNullOrEmpty(help)) liba.Add("helpstring(\"" + help + "\")");
            if (cnt != 0) liba.Add("helpcontext(" + cnt.PaddedHex() + ")");

            cnt = 0;
            liba.ForEach(x => ih.AppendLine("  " + x + (++cnt == liba.Count ? "" : ",")));
            ih.AppendLine("]");
            ih.AppendLine("module " + _name + " {");
            using (new IDLHelperTab(ih))
            {
                Children.ForEach(x => x.BuildIDLInto(ih));
            }
            ih.AppendLine("};");
        }
    }

    class OWChildrenIndirect : ITlibNode
    {
        ITlibNode _parent;
        int _imageidx;
        String _name;
        dlgCreateChildren _genChildren;
        public OWChildrenIndirect(ITlibNode parent, String name, int imageidx, dlgCreateChildren genchildren)
        {
            _parent = parent;
            _name = name;
            _imageidx = imageidx;
            _genChildren = genchildren;
        }
        public override String Name
        {
            get
            {
                return _name;
            }
        }
        public override string ShortName
        {
            get { return _name; }
        }
        public override string ObjectName
        {
            get { return null; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return true;
        }
        public override int ImageIndex
        {
            get { return _imageidx; }
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            return _genChildren();
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            Children.ForEach(x => x.BuildIDLInto(ih));
        }
    }


    class OWModuleConst : ITlibNode
    {
        ITlibNode _parent;
        String _name;
        VarDesc _vd;
        ITypeInfo _ti;
        Object _val;
        int _idx;
        public OWModuleConst(ITlibNode parent, ITypeInfo ti, VarDesc vd, int idx)
        {
            _parent = parent;
            _vd = vd;
            _ti = ti;
            IDLGrabber ig = new IDLGrabber();
            _vd.elemDescVar.tdesc.ComTypeNameAsString(_ti, ig);
            _name = ig.Value + " " + ti.GetDocumentationById(vd.memid);
            _val = vd.varValue;
            if (_val == null) _val = "";
            if (_val.GetType() == typeof(String))
            {
                _val = (_val as String).ReEscape();
            }
            _idx = idx;
        }
        public override String Name
        {
            get
            {
                return "const " + _name + " = " + _val;
            }
        }
        public override string ShortName
        {
            get { return _name; }
        }
        public override string ObjectName
        {
            get { return null;  }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return false;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_const; }
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            return res;
        }
        String negStr(int x)
        {
            return (x < 0) ? ("0x" + x.ToString("X")) : x.ToString();
        }
        public void BuildIDLInto(IDLFormatter ih, bool embedded, bool islast)
        {
            String desc = "";
            //int cnt = 0;
            //String help = _ti.GetHelpDocumentationById(_idx, out cnt);
            //List<String> props = new List<string>();
            //AddHelpStringAndContext(props, help, cnt);
            //if (props.Count > 0)
            //{
            //    desc += "[" + String.Join(",", props.ToArray()) + "] ";
            //}
            desc += _val.GetType() == typeof(int) ? negStr((int)_val) : _val.ToString();
            ih.AppendLine("const " + _name + " = " + desc + (embedded ? (islast ? "" : ",") : ";"));
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            BuildIDLInto(ih, false, false);
        }
    }


    class OWRecord : ITlibNode
    {
        ITlibNode _parent;
        String _name;
        ITypeInfo _ti;
        TypeAttr _ta;
        public OWRecord(ITlibNode parent, ITypeInfo ti, TypeAttr ta)
        {
            _parent = parent;
            _ti = ti;
            _ta = ta;
            _name = _ti.GetName();
        }
        public override String Name
        {
            get
            {
                return "typedef struct " + _name;
            }
        }
        public override string ShortName
        {
            get { return _name;  }
        }
        public override string ObjectName
        {
            get { return _name + "#s"; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return true;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_struct; }
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            for (int x = 0; x < _ta.cVars; ++x)
            {
                VarDesc vd = new VarDesc(_ti, x);
                res.Add(new OWRecordMember(this, _ti, vd));
            }
            return res;
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            ih.AppendLine("typedef struct tag" + _name + " {");
            using (new IDLHelperTab(ih))
            {
                Children.ForEach( x => x.BuildIDLInto(ih) );
            }
            ih.AppendLine("} " + _name + ";");
        }
    }

    class OWRecordMember : ITlibNode
    {
        ITlibNode _parent;
        String _type;
        String _name;
        VarDesc _vd;
        ITypeInfo _ti;

        public OWRecordMember(ITlibNode parent, ITypeInfo ti, VarDesc vd)
        {
            _parent = parent;
            _name = ti.GetDocumentationById(vd.memid);
            _vd = vd;
            _ti = ti;
            IDLGrabber ig = new IDLGrabber();
            _vd.elemDescVar.tdesc.ComTypeNameAsString(_ti, ig);
            _type = ig.Value;
        }
        public override String Name
        {
            get
            {            
                return _type + " " + _name;
            }
        }
        public override string ShortName
        {
            get { return _name; }
        }
        public override string ObjectName
        {
            get { return null; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return false;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_strucmem; }
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            return res;
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            ih.AppendLine(this.Name + ";");
        }
    }

    class OWTypeDef : ITlibNode
    {
        ITlibNode _parent;
        ITypeInfo _ti;
        TypeAttr _ta;
        String _name;

        public OWTypeDef(ITlibNode parent, ITypeInfo ti, TypeAttr ta)
        {
            _parent = parent;
            _ta = ta;
            _ti = ti;

            ITypeInfo oti;
            _ti.GetRefTypeInfo(_ta.tdescAlias.hreftype, out oti);
            _name = oti.GetName() + " " + ti.GetName();        
        }
        public override String Name
        {
            get
            {
                return "typedef " + _name;
            }
        }
        public override string ShortName
        {
            get { return _name; }
        }
        public override string ObjectName
        {
            get { return _name + "#i"; }
        }
        public override bool DisplayAtTLBLevel(ICollection<String> interfaceNames)
        {
            return true;
        }
        public override int ImageIndex
        {
            get { return (int)ImageIndices.idx_typedef; }
        }
        public override ITlibNode Parent { get { return _parent; } }
        public override List<ITlibNode> GenChildren()
        {
            List<ITlibNode> res = new List<ITlibNode>();
            ITypeInfo oti;
            _ti.GetRefTypeInfo(_ta.tdescAlias.hreftype, out oti);
            CommonBuildTlibNode(this, oti, false, false, res);
            return res;
        }
        public override void BuildIDLInto(IDLFormatter ih)
        {
            ih.AppendLine("typedef [public] " + _name + ";");
        }
    }

}

