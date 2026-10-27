from lxml import etree

from slideforge.omml import MATHML, height_pt, mathml


def rows(latex: str) -> list[str]:
    """The text of each mrow the adjustments added, in document order."""
    root = mathml(latex)
    return [
        "".join(row.itertext())
        for row in root.iter(f"{{{MATHML}}}mrow")
        if row.getparent().tag != f"{{{MATHML}}}math"
    ]


def test_function_name_takes_its_argument():
    assert "sin⁡x" in rows(r"e^{iS\sin x}")


def test_sum_operand_runs_to_the_next_plus():
    assert "Jn(S)2" in rows(r"\sum_n J_n(S)^2 + 1")


def test_sum_operand_keeps_minus_inside_brackets():
    assert "Jn(SA(t−z)) dz" in rows(r"\int_0^L J_n(S A(t-z))\,dz = 1")


def test_single_operand_is_not_wrapped():
    root = mathml(r"\sum_n a_n")
    assert len(list(root.iter(f"{{{MATHML}}}mrow"))) == 1


def test_mathml_is_display():
    assert mathml("x").get("display") == "block"
    assert etree.QName(mathml("x")).localname == "math"


def test_capital_greek_and_nabla_are_upright():
    root = mathml(r"\nabla\cdot\Theta + \theta")
    upright = [n.text for n in root.iter() if n.get("mathvariant") == "normal"]
    assert upright == ["∇", "Θ"]


def test_spaces_survive_as_unicode_spaces():
    texts = [n.text for n in mathml(r"a\qquad b\,c").iter(f"{{{MATHML}}}mtext")]
    assert texts == ["  ", " "]


def test_negative_space_is_dropped():
    assert not list(mathml(r"a\!b").iter(f"{{{MATHML}}}mspace", f"{{{MATHML}}}mtext"))


def test_height_grows_with_fractions_and_limits():
    plain = height_pt("x = 1", 24)
    fraction = height_pt(r"x = \frac{a}{b}", 24)
    limits = height_pt(r"\sum_{n=0}^{\infty} x^n", 24)
    assert plain < fraction < limits
