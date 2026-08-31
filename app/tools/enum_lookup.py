"""Read-only enum constant lookup via pythoncom typelib introspection.
Does NOT use gencache/EnsureDispatch/EnsureModule — safe to freely mix with GetActiveObject."""
import sys
import win32com.client


def get_typelib(app):
    ti = app._oleobj_.GetTypeInfo()
    tlib, idx = ti.GetContainingTypeLib()
    return tlib


def list_enum(tlib, enum_type_name):
    n = tlib.GetTypeInfoCount()
    for i in range(n):
        if tlib.GetTypeInfoType(i) != 0:  # TKIND_ENUM
            continue
        name = tlib.GetDocumentation(i)[0]
        if name != enum_type_name:
            continue
        ti = tlib.GetTypeInfo(i)
        ta = ti.GetTypeAttr()
        cVars = ta[7]
        out = {}
        for v in range(cVars):
            vd = ti.GetVarDesc(v)
            varnames = ti.GetNames(vd[0])
            value = vd[1]
            out[varnames[0]] = value
        return out
    return None


if __name__ == "__main__":
    app = win32com.client.GetActiveObject("Inventor.Application")
    tlib = get_typelib(app)
    for enum_name in sys.argv[1:]:
        result = list_enum(tlib, enum_name)
        print(f"=== {enum_name} ===")
        if result is None:
            print("  NOT FOUND")
        else:
            for k, v in sorted(result.items(), key=lambda kv: kv[1]):
                print(f"  {k} = {v}")
