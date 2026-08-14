/* @ds-bundle: {"format":3,"namespace":"CHAIDesignSystem_eb475a","components":[{"name":"Badge","sourcePath":"components/core/Badge.jsx"},{"name":"Button","sourcePath":"components/core/Button.jsx"},{"name":"Card","sourcePath":"components/core/Card.jsx"},{"name":"Callout","sourcePath":"components/feedback/Callout.jsx"},{"name":"Stat","sourcePath":"components/feedback/Stat.jsx"}],"sourceHashes":{"components/core/Badge.jsx":"4d9299a2a2b7","components/core/Button.jsx":"ac492b03df09","components/core/Card.jsx":"d9d17e5b1157","components/feedback/Callout.jsx":"dd0b923a07c8","components/feedback/Stat.jsx":"d43311208afb","ui_kits/slides/Slides.jsx":"c9ea201ecf59","ui_kits/slides/SlidesB.jsx":"d0a1b0213b30"},"inlinedExternals":[],"unexposedExports":[]} */

(() => {

const __ds_ns = (window.CHAIDesignSystem_eb475a = window.CHAIDesignSystem_eb475a || {});

const __ds_scope = {};

(__ds_ns.__errors = __ds_ns.__errors || []);

// components/core/Badge.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
const TONES = {
  blue: {
    bg: "var(--chai-light-blue)",
    fg: "var(--chai-dark-blue)"
  },
  teal: {
    bg: "var(--chai-teal)",
    fg: "#08312c"
  },
  green: {
    bg: "var(--chai-green)",
    fg: "#06492b"
  },
  gold: {
    bg: "var(--chai-gold)",
    fg: "#3d2c00"
  },
  red: {
    bg: "var(--chai-dark-red)",
    fg: "#ffffff"
  },
  grey: {
    bg: "var(--chai-light-grey)",
    fg: "var(--chai-dark-grey)"
  }
};

/** CHAI Badge / status pill. */
function Badge({
  children,
  tone = "blue",
  style = {},
  ...rest
}) {
  const t = TONES[tone] || TONES.blue;
  return /*#__PURE__*/React.createElement("span", _extends({
    style: {
      display: "inline-flex",
      alignItems: "center",
      gap: "6px",
      background: t.bg,
      color: t.fg,
      fontFamily: "var(--font-base)",
      fontSize: "var(--text-xs)",
      fontWeight: "var(--weight-bold)",
      letterSpacing: "var(--tracking-wide)",
      textTransform: "uppercase",
      padding: "4px 10px",
      borderRadius: "var(--radius-pill)",
      lineHeight: 1.2,
      ...style
    }
  }, rest), children);
}
Object.assign(__ds_scope, { Badge });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/core/Badge.jsx", error: String((e && e.message) || e) }); }

// components/core/Button.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * CHAI Button — primary action uses Dark Blue, accent uses Turquoise.
 */
function Button({
  children,
  variant = "primary",
  size = "md",
  disabled = false,
  type = "button",
  onClick,
  style = {},
  ...rest
}) {
  const sizes = {
    sm: {
      padding: "6px 14px",
      fontSize: "var(--text-sm)"
    },
    md: {
      padding: "10px 20px",
      fontSize: "var(--text-base)"
    },
    lg: {
      padding: "13px 28px",
      fontSize: "var(--text-lg)"
    }
  };
  const variants = {
    primary: {
      background: "var(--color-primary)",
      color: "var(--color-text-inverse)",
      border: "var(--border-width) solid var(--color-primary)"
    },
    accent: {
      background: "var(--color-accent)",
      color: "var(--color-text-inverse)",
      border: "var(--border-width) solid var(--color-accent)"
    },
    secondary: {
      background: "transparent",
      color: "var(--color-primary)",
      border: "var(--border-width-strong) solid var(--color-primary)"
    },
    ghost: {
      background: "transparent",
      color: "var(--color-primary)",
      border: "var(--border-width-strong) solid transparent"
    }
  };
  return /*#__PURE__*/React.createElement("button", _extends({
    type: type,
    disabled: disabled,
    onClick: onClick,
    style: {
      display: "inline-flex",
      alignItems: "center",
      justifyContent: "center",
      gap: "8px",
      fontFamily: "var(--font-base)",
      fontWeight: "var(--weight-bold)",
      lineHeight: 1,
      borderRadius: "var(--radius-md)",
      cursor: disabled ? "not-allowed" : "pointer",
      opacity: disabled ? 0.45 : 1,
      transition: "background var(--duration) var(--ease), border-color var(--duration) var(--ease)",
      ...sizes[size],
      ...variants[variant],
      ...style
    },
    onMouseEnter: e => {
      if (disabled) return;
      if (variant === "primary") e.currentTarget.style.background = "var(--color-primary-hover)";else if (variant === "accent") e.currentTarget.style.background = "var(--color-accent-hover)";else e.currentTarget.style.background = "var(--chai-light-blue)";
    },
    onMouseLeave: e => {
      e.currentTarget.style.background = variants[variant].background;
    }
  }, rest), children);
}
Object.assign(__ds_scope, { Button });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/core/Button.jsx", error: String((e && e.message) || e) }); }

// components/core/Card.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/** CHAI Card — clean white surface, soft shadow, restrained rounding. */
function Card({
  children,
  padded = true,
  accent = false,
  style = {},
  ...rest
}) {
  return /*#__PURE__*/React.createElement("div", _extends({
    style: {
      background: "var(--color-surface)",
      border: "var(--border-width) solid var(--color-border)",
      borderTop: accent ? "4px solid var(--color-accent)" : undefined,
      borderRadius: "var(--radius-lg)",
      boxShadow: "var(--shadow-md)",
      padding: padded ? "var(--space-5)" : 0,
      fontFamily: "var(--font-base)",
      color: "var(--color-text)",
      ...style
    }
  }, rest), children);
}
Object.assign(__ds_scope, { Card });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/core/Card.jsx", error: String((e && e.message) || e) }); }

// components/feedback/Callout.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
const TONES = {
  blue: {
    bg: "var(--chai-light-blue)",
    fg: "var(--chai-dark-blue)",
    bar: "var(--chai-dark-blue)"
  },
  gold: {
    bg: "var(--chai-light-gold)",
    fg: "#3d2c00",
    bar: "var(--chai-gold)"
  },
  grey: {
    bg: "var(--chai-light-grey)",
    fg: "var(--chai-dark-grey)",
    bar: "var(--chai-grey)"
  }
};

/**
 * CHAI Callout box — the signature highlight block.
 * Light Blue background with Dark Blue text (per Identity Guide).
 */
function Callout({
  children,
  title,
  tone = "blue",
  style = {},
  ...rest
}) {
  const t = TONES[tone] || TONES.blue;
  return /*#__PURE__*/React.createElement("div", _extends({
    style: {
      background: t.bg,
      color: t.fg,
      borderLeft: `4px solid ${t.bar}`,
      borderRadius: "var(--radius-md)",
      padding: "var(--space-4) var(--space-5)",
      fontFamily: "var(--font-base)",
      fontSize: "var(--text-base)",
      lineHeight: "var(--leading-normal)",
      ...style
    }
  }, rest), title && /*#__PURE__*/React.createElement("div", {
    style: {
      fontWeight: "var(--weight-bold)",
      fontSize: "var(--text-lg)",
      marginBottom: "var(--space-2)"
    }
  }, title), children);
}
Object.assign(__ds_scope, { Callout });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/feedback/Callout.jsx", error: String((e && e.message) || e) }); }

// components/feedback/Stat.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/** CHAI Stat — large impact figure with label. Used for program metrics. */
function Stat({
  value,
  label,
  sublabel,
  tone = "blue",
  align = "left",
  style = {},
  ...rest
}) {
  const colors = {
    blue: "var(--chai-dark-blue)",
    turquoise: "var(--chai-turquoise)",
    green: "var(--chai-green)",
    gold: "var(--chai-gold)"
  };
  return /*#__PURE__*/React.createElement("div", _extends({
    style: {
      fontFamily: "var(--font-base)",
      textAlign: align,
      ...style
    }
  }, rest), /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "var(--font-display)",
      fontSize: "var(--text-4xl)",
      fontWeight: "var(--weight-bold)",
      lineHeight: 1,
      color: colors[tone] || colors.blue,
      letterSpacing: "var(--tracking-tight)"
    }
  }, value), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: "var(--space-2)",
      fontSize: "var(--text-base)",
      fontWeight: "var(--weight-bold)",
      color: "var(--color-text)"
    }
  }, label), sublabel && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: "2px",
      fontSize: "var(--text-sm)",
      color: "var(--color-text-muted)"
    }
  }, sublabel));
}
Object.assign(__ds_scope, { Stat });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/feedback/Stat.jsx", error: String((e && e.message) || e) }); }

// ui_kits/slides/Slides.jsx
try { (() => {
/* global React */
// Shared slide frame: fixed 1280x720, scales to fit its container.
function Slide({
  children,
  bg = "#fff",
  pad = 64,
  style = {}
}) {
  return /*#__PURE__*/React.createElement("div", {
    style: {
      width: 1280,
      height: 720,
      background: bg,
      position: "relative",
      fontFamily: "var(--font-base)",
      overflow: "hidden",
      padding: pad,
      boxSizing: "border-box",
      ...style
    }
  }, children);
}
function Logo({
  variant = "blue",
  h = 40,
  style = {}
}) {
  return /*#__PURE__*/React.createElement("img", {
    src: `../../assets/chai-logo-${variant}.svg`,
    alt: "CHAI",
    style: {
      height: h,
      ...style
    }
  });
}
window.Slide = Slide;
window.Logo = Logo;

// ---- Title slide ----
function TitleSlide() {
  return /*#__PURE__*/React.createElement(window.Slide, {
    bg: "var(--chai-dark-blue)",
    pad: 0
  }, /*#__PURE__*/React.createElement("img", {
    src: "../../assets/photo-pharmacy.jpeg",
    alt: "",
    style: {
      position: "absolute",
      inset: 0,
      width: "100%",
      height: "100%",
      objectFit: "cover",
      opacity: 0.28
    }
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      position: "absolute",
      inset: 0,
      background: "linear-gradient(105deg, rgba(0,43,83,0.95) 30%, rgba(0,43,83,0.55) 100%)"
    }
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      position: "relative",
      height: "100%",
      padding: 80,
      boxSizing: "border-box",
      display: "flex",
      flexDirection: "column",
      justifyContent: "space-between",
      color: "#fff"
    }
  }, /*#__PURE__*/React.createElement(window.Logo, {
    variant: "white",
    h: 52
  }), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 18,
      fontWeight: 700,
      letterSpacing: ".06em",
      textTransform: "uppercase",
      color: "var(--chai-teal)",
      marginBottom: 20
    }
  }, "Program Review 2025"), /*#__PURE__*/React.createElement("h1", {
    style: {
      fontFamily: "var(--font-base)",
      fontSize: 72,
      lineHeight: 1.05,
      fontWeight: 700,
      margin: 0,
      maxWidth: 920,
      letterSpacing: "-.01em"
    }
  }, "Strengthening health systems that succeed without us")), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 16,
      color: "rgba(255,255,255,0.75)"
    }
  }, "Global Communications Team \xA0\xB7\xA0 clintonhealthaccess.org")));
}
window.TitleSlide = TitleSlide;

// ---- Section header ----
function SectionSlide() {
  return /*#__PURE__*/React.createElement(window.Slide, {
    bg: "var(--chai-light-blue)"
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      height: "100%",
      display: "flex",
      flexDirection: "column",
      justifyContent: "center"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 20,
      fontWeight: 700,
      color: "var(--chai-turquoise)",
      letterSpacing: ".04em",
      textTransform: "uppercase",
      marginBottom: 16
    }
  }, "Section 02"), /*#__PURE__*/React.createElement("h2", {
    style: {
      fontFamily: "var(--font-base)",
      fontSize: 64,
      fontWeight: 700,
      color: "var(--chai-dark-blue)",
      margin: 0,
      maxWidth: 900
    }
  }, "Our impact across program countries"), /*#__PURE__*/React.createElement("div", {
    style: {
      width: 96,
      height: 6,
      background: "var(--chai-turquoise)",
      marginTop: 32,
      borderRadius: 3
    }
  })));
}
window.SectionSlide = SectionSlide;
})(); } catch (e) { __ds_ns.__errors.push({ path: "ui_kits/slides/Slides.jsx", error: String((e && e.message) || e) }); }

// ui_kits/slides/SlidesB.jsx
try { (() => {
/* global React */
// ---- Stats slide ----
function StatsSlide() {
  const {
    Stat
  } = window.CHAIDesignSystem_eb475a;
  const items = [{
    value: "35+",
    label: "Countries",
    sublabel: "where we work"
  }, {
    value: "85%",
    label: "Staff",
    sublabel: "in program countries"
  }, {
    value: "150M+",
    label: "People reached",
    sublabel: "through programs"
  }, {
    value: "20+",
    label: "Years",
    sublabel: "of partnership"
  }];
  return /*#__PURE__*/React.createElement(window.Slide, null, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
      marginBottom: 48
    }
  }, /*#__PURE__*/React.createElement("h2", {
    style: {
      fontFamily: "var(--font-base)",
      fontSize: 44,
      fontWeight: 700,
      color: "var(--chai-dark-blue)",
      margin: 0
    }
  }, "By the numbers"), /*#__PURE__*/React.createElement(window.Logo, {
    h: 36
  })), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "repeat(2,1fr)",
      gap: 40
    }
  }, items.map((it, i) => /*#__PURE__*/React.createElement("div", {
    key: i,
    style: {
      background: "var(--chai-light-grey)",
      borderRadius: 10,
      padding: 36,
      borderLeft: "6px solid var(--chai-turquoise)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "var(--font-base)",
      fontSize: 64,
      fontWeight: 700,
      color: "var(--chai-dark-blue)",
      lineHeight: 1
    }
  }, it.value), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 22,
      fontWeight: 700,
      marginTop: 12,
      color: "var(--color-text)"
    }
  }, it.label), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 17,
      color: "var(--color-text-muted)",
      marginTop: 4
    }
  }, it.sublabel)))));
}
window.StatsSlide = StatsSlide;

// ---- Quote slide ----
function QuoteSlide() {
  return /*#__PURE__*/React.createElement(window.Slide, {
    bg: "var(--chai-dark-blue)"
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      height: "100%",
      display: "flex",
      flexDirection: "column",
      justifyContent: "center",
      color: "#fff",
      maxWidth: 1000
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "var(--font-base)",
      fontSize: 120,
      lineHeight: 0.6,
      color: "var(--chai-teal)"
    }
  }, "\u201C"), /*#__PURE__*/React.createElement("blockquote", {
    style: {
      fontFamily: "var(--font-base)",
      fontSize: 44,
      fontWeight: 700,
      lineHeight: 1.25,
      margin: "16px 0 32px"
    }
  }, "We communicate our unique approach to accomplishing our mission with one voice, built on consistent language and visuals."), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 18,
      color: "var(--chai-teal)",
      fontWeight: 700
    }
  }, "CHAI Identity Guide")));
}
window.QuoteSlide = QuoteSlide;

// ---- Content / two-column slide ----
function ContentSlide() {
  const {
    Callout,
    Badge
  } = window.CHAIDesignSystem_eb475a;
  return /*#__PURE__*/React.createElement(window.Slide, null, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
      marginBottom: 36
    }
  }, /*#__PURE__*/React.createElement("h2", {
    style: {
      fontFamily: "var(--font-base)",
      fontSize: 44,
      fontWeight: 700,
      color: "var(--chai-dark-blue)",
      margin: 0
    }
  }, "Our approach"), /*#__PURE__*/React.createElement(window.Logo, {
    h: 36
  })), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "1.3fr 1fr",
      gap: 48,
      alignItems: "start"
    }
  }, /*#__PURE__*/React.createElement("div", null, [["Partner with governments", "We embed alongside ministries of health to strengthen systems from within."], ["Lower the cost of care", "We negotiate price reductions so treatments reach more people."], ["Build for sustainability", "We create capacity that endures long after our involvement ends."]].map(([h, b], i) => /*#__PURE__*/React.createElement("div", {
    key: i,
    style: {
      display: "flex",
      gap: 18,
      marginBottom: 28
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      flex: "none",
      width: 40,
      height: 40,
      borderRadius: "50%",
      background: "var(--chai-turquoise)",
      color: "#fff",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontWeight: 700,
      fontSize: 18
    }
  }, i + 1), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 22,
      fontWeight: 700,
      color: "var(--color-text-strong)"
    }
  }, h), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 17,
      color: "var(--color-text-muted)",
      lineHeight: 1.5,
      marginTop: 4
    }
  }, b))))), /*#__PURE__*/React.createElement(Callout, {
    title: "The result"
  }, "High-quality health systems that can succeed without our assistance.")));
}
window.ContentSlide = ContentSlide;
})(); } catch (e) { __ds_ns.__errors.push({ path: "ui_kits/slides/SlidesB.jsx", error: String((e && e.message) || e) }); }

__ds_ns.Badge = __ds_scope.Badge;

__ds_ns.Button = __ds_scope.Button;

__ds_ns.Card = __ds_scope.Card;

__ds_ns.Callout = __ds_scope.Callout;

__ds_ns.Stat = __ds_scope.Stat;

})();
