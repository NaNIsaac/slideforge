from slideforge.omml import omml
from slideforge.xml import q

from .conftest import needs_office_math

pytestmark = needs_office_math


def texts(element) -> str:
    return "".join(t.text or "" for t in element.iter(q("m:t")))


def test_bracket_after_subscript_stays_in_the_sum():
    math = omml(r"\sum_n J_n(S)^2 + 1")
    operand = math.find(f"{q('m:nary')}/{q('m:e')}")
    assert texts(operand) == "Jn(S)2"


def test_integral_operand_includes_the_differential():
    math = omml(r"\int_0^L J_n(S\,A(t-z/v_s))\,dz")
    assert texts(math.find(f"{q('m:nary')}/{q('m:e')}")).endswith("dz")


def test_sin_is_a_function():
    func = omml(r"\sin x").find(f".//{q('m:func')}")
    assert func is not None
    assert texts(func.find(q("m:fName"))) == "sin"
    assert texts(func.find(q("m:e"))) == "x"


def test_bold_letters_are_plain_runs_with_bold_style():
    math = omml(r"I(\mathbf{Q})")
    bold = [r for r in math.iter(q("m:r")) if r.find(f"{q('m:rPr')}/{q('m:sty')}") is not None]
    assert [texts(r) for r in bold] == ["Q"]


def test_runs_carry_font_and_size():
    rpr = omml("x", font="Fira Math", size_pt=20).find(f".//{q('a:rPr')}")
    assert rpr.get("sz") == "2000"
    assert rpr.find(q("a:latin")).get("typeface") == "Fira Math"


def test_root_is_one_equation():
    assert omml("x=1").tag == q("m:oMath")
