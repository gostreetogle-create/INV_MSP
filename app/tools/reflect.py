"""Read-only COM type introspection via pythoncom — does NOT touch gencache, safe to mix
with plain GetActiveObject dynamic dispatch."""
import sys
import pythoncom
import win32com.client


def reflect(obj, name="object"):
    ti = obj._oleobj_.GetTypeInfo()
    ta = ti.GetTypeAttr()
    print(f"--- {name}: raw attr={ta}")
    cFuncs = ta[6]
    for i in range(cFuncs):
        try:
            fd = ti.GetFuncDesc(i)
        except Exception as e:
            print("  funcdesc error", i, e)
            continue
        memid = fd[0]
        try:
            names = ti.GetNames(memid)
        except Exception as e:
            names = [f"<id {memid}>"]
        fname = names[0]
        argnames = names[1:]
        params = fd[2]
        opt = fd[6]
        print(f"  [{i}] {fname}({', '.join(argnames)}) nParams={len(params)} nOptional={opt} invkind={fd[4]}")


def reflect_one(obj, method_name):
    ti = obj._oleobj_.GetTypeInfo()
    ta = ti.GetTypeAttr()
    for i in range(ta[6]):
        fd = ti.GetFuncDesc(i)
        names = ti.GetNames(fd[0])
        if names[0] == method_name:
            print(f"{method_name} raw funcdesc params:")
            for p in fd[2]:
                tdesc = p[0]
                print("  ", p)
                if isinstance(tdesc, tuple) and tdesc[0] == 26 and isinstance(tdesc[1], tuple) and tdesc[1][0] == 29:
                    href = tdesc[1][1]
                    refti = ti.GetRefTypeInfo(href)
                    refattr = refti.GetTypeAttr()
                    print("     -> ref type name:", refti.GetDocumentation(-1)[0], "typekind:", refattr[5])
            return
    print("not found:", method_name)


if __name__ == "__main__":
    app = win32com.client.GetActiveObject("Inventor.Application")
    tpl = r"C:\Users\Public\Documents\Autodesk\Inventor 2027\Templates\ru-RU\Metric\Sheet Metal (mm).ipt"
    doc = app.Documents.Add(12290, tpl, True)
    compDef = doc.ComponentDefinition
    reflect_one(compDef.Features.FlangeFeatures, "CreateFlangeDefinition")
    doc.Close(False)
