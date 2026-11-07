//! JavaScript value semantics for the port of the TypeScript save.
//!
//! The TypeScript save works on plain JSON-like objects: shared by reference, compared by
//! identity, keyed in insertion order, with `undefined` distinct from `null` and every number
//! an `f64`. [`Js`] reproduces that so the port can follow the source line by line.

use std::cell::{Ref, RefCell, RefMut};
use std::collections::HashMap;
use std::fmt;
use std::rc::{Rc, Weak};

use indexmap::IndexMap;
use serde::de::{self, Deserialize, Deserializer, MapAccess, SeqAccess, Visitor};
use serde::ser::{Serialize, SerializeMap, SerializeSeq, Serializer};

/// The hasher of every object and map: foldhash, as docx-edit uses.
pub type Hasher = foldhash::fast::RandomState;
pub type Props = IndexMap<String, Js, Hasher>;
pub type StrMap = IndexMap<String, Js, Hasher>;

/// A JavaScript value. Objects, arrays and maps are shared references, as in JS.
#[derive(Clone, Default)]
pub enum Js {
    #[default]
    Undefined,
    Null,
    Bool(bool),
    Num(f64),
    Str(Rc<str>),
    Arr(Arr),
    Obj(Obj),
    /// A JS `Map` with string keys.
    Map(JsMap),
    /// An `ArrayBuffer` or `Uint8Array`.
    Bytes(Rc<Vec<u8>>),
}

#[derive(Clone)]
pub struct Obj(pub Rc<RefCell<Props>>);
#[derive(Clone)]
pub struct Arr(pub Rc<RefCell<Vec<Js>>>);
#[derive(Clone)]
pub struct JsMap(pub Rc<RefCell<StrMap>>);

/// A thrown error.
#[derive(Debug, Clone, PartialEq)]
pub struct SaveError {
    /// `"unsupported"` for a `DocxSectionError('unsupported', …)`, otherwise empty.
    pub code: &'static str,
    pub message: String,
}

impl SaveError {
    pub fn new(message: impl Into<String>) -> Self {
        Self {
            code: "",
            message: message.into(),
        }
    }
    pub fn section_unsupported(message: impl Into<String>) -> Self {
        Self {
            code: "unsupported",
            message: message.into(),
        }
    }
}

impl fmt::Display for SaveError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.message)
    }
}

impl std::error::Error for SaveError {}

pub type Result<T> = std::result::Result<T, SaveError>;

/// `throw new Error(…)`.
#[macro_export]
macro_rules! throw {
    ($($arg:tt)*) => {
        return Err($crate::js::SaveError::new(format!($($arg)*)))
    };
}

/// An object literal: `obj! { "type" => "run", "content" => content }`. Values go through
/// `Js::from`; an `Js::Undefined` value is kept as a property, as `{ a: undefined }` is.
#[macro_export]
macro_rules! obj {
    () => { $crate::js::Js::new_obj() };
    ($($key:expr => $value:expr),+ $(,)?) => {{
        let o = $crate::js::Js::new_obj();
        $( o.set($key, $crate::js::Js::from($value)); )+
        o
    }};
}

/// An array literal: `arr![a, b]`.
#[macro_export]
macro_rules! arr {
    () => { $crate::js::Js::new_arr(Vec::new()) };
    ($($value:expr),+ $(,)?) => {
        $crate::js::Js::new_arr(vec![$($crate::js::Js::from($value)),+])
    };
}

impl From<bool> for Js {
    fn from(v: bool) -> Self {
        Js::Bool(v)
    }
}
impl From<f64> for Js {
    fn from(v: f64) -> Self {
        Js::Num(v)
    }
}
impl From<i64> for Js {
    fn from(v: i64) -> Self {
        Js::Num(v as f64)
    }
}
impl From<i32> for Js {
    fn from(v: i32) -> Self {
        Js::Num(v as f64)
    }
}
impl From<u32> for Js {
    fn from(v: u32) -> Self {
        Js::Num(v as f64)
    }
}
impl From<usize> for Js {
    fn from(v: usize) -> Self {
        Js::Num(v as f64)
    }
}
impl From<&str> for Js {
    fn from(v: &str) -> Self {
        Js::Str(Rc::from(v))
    }
}
impl From<String> for Js {
    fn from(v: String) -> Self {
        Js::Str(Rc::from(v))
    }
}
impl From<&String> for Js {
    fn from(v: &String) -> Self {
        Js::Str(Rc::from(v.as_str()))
    }
}
impl From<Rc<str>> for Js {
    fn from(v: Rc<str>) -> Self {
        Js::Str(v)
    }
}
impl From<Vec<Js>> for Js {
    fn from(v: Vec<Js>) -> Self {
        Js::new_arr(v)
    }
}
impl From<&Js> for Js {
    fn from(v: &Js) -> Self {
        v.clone()
    }
}
impl<T: Into<Js>> From<Option<T>> for Js {
    fn from(v: Option<T>) -> Self {
        v.map_or(Js::Undefined, Into::into)
    }
}

impl Js {
    pub fn new_obj() -> Js {
        Js::Obj(Obj(Rc::new(RefCell::new(Props::default()))))
    }
    pub fn obj_from(props: Props) -> Js {
        Js::Obj(Obj(Rc::new(RefCell::new(props))))
    }
    pub fn new_arr(items: Vec<Js>) -> Js {
        Js::Arr(Arr(Rc::new(RefCell::new(items))))
    }
    pub fn new_map() -> Js {
        Js::Map(JsMap(Rc::new(RefCell::new(StrMap::default()))))
    }
    pub fn map_from(entries: impl IntoIterator<Item = (String, Js)>) -> Js {
        Js::Map(JsMap(Rc::new(RefCell::new(entries.into_iter().collect()))))
    }
    pub fn bytes(bytes: Vec<u8>) -> Js {
        Js::Bytes(Rc::new(bytes))
    }

    // ----- type tests -----

    pub fn is_undefined(&self) -> bool {
        matches!(self, Js::Undefined)
    }
    pub fn is_null(&self) -> bool {
        matches!(self, Js::Null)
    }
    /// `x == null`.
    pub fn is_nullish(&self) -> bool {
        matches!(self, Js::Undefined | Js::Null)
    }
    /// A plain object (not an array, map or null).
    pub fn is_obj(&self) -> bool {
        matches!(self, Js::Obj(_))
    }
    /// `typeof x === 'object' && x !== null`.
    pub fn is_object_like(&self) -> bool {
        matches!(self, Js::Obj(_) | Js::Arr(_) | Js::Map(_) | Js::Bytes(_))
    }
    pub fn is_arr(&self) -> bool {
        matches!(self, Js::Arr(_))
    }
    pub fn is_map(&self) -> bool {
        matches!(self, Js::Map(_))
    }
    pub fn is_str(&self) -> bool {
        matches!(self, Js::Str(_))
    }
    pub fn is_num(&self) -> bool {
        matches!(self, Js::Num(_))
    }
    pub fn is_bool(&self) -> bool {
        matches!(self, Js::Bool(_))
    }
    pub fn is_bytes(&self) -> bool {
        matches!(self, Js::Bytes(_))
    }
    /// `typeof x`.
    pub fn type_of(&self) -> &'static str {
        match self {
            Js::Undefined => "undefined",
            Js::Bool(_) => "boolean",
            Js::Num(_) => "number",
            Js::Str(_) => "string",
            _ => "object",
        }
    }
    /// JS truthiness.
    pub fn truthy(&self) -> bool {
        match self {
            Js::Undefined | Js::Null => false,
            Js::Bool(b) => *b,
            Js::Num(n) => *n != 0.0 && !n.is_nan(),
            Js::Str(s) => !s.is_empty(),
            _ => true,
        }
    }
    /// `x ?? fallback`.
    pub fn or_else(self, fallback: impl FnOnce() -> Js) -> Js {
        if self.is_nullish() { fallback() } else { self }
    }
    /// `x ?? fallback` with a ready fallback.
    pub fn nullish_or(self, fallback: Js) -> Js {
        if self.is_nullish() { fallback } else { self }
    }
    /// `x || fallback`.
    pub fn or(self, fallback: Js) -> Js {
        if self.truthy() { self } else { fallback }
    }
    /// The value when it is defined: `x === undefined ? None : Some(x)`.
    pub fn defined(self) -> Option<Js> {
        if self.is_undefined() {
            None
        } else {
            Some(self)
        }
    }
    /// The value when it is not nullish.
    pub fn some(self) -> Option<Js> {
        if self.is_nullish() { None } else { Some(self) }
    }

    // ----- primitives -----

    pub fn as_str(&self) -> Option<&str> {
        match self {
            Js::Str(s) => Some(s),
            _ => None,
        }
    }
    /// The string, or `""` for any other value.
    pub fn str_or_empty(&self) -> &str {
        self.as_str().unwrap_or("")
    }
    pub fn as_rc_str(&self) -> Option<Rc<str>> {
        match self {
            Js::Str(s) => Some(s.clone()),
            _ => None,
        }
    }
    pub fn as_f64(&self) -> Option<f64> {
        match self {
            Js::Num(n) => Some(*n),
            _ => None,
        }
    }
    pub fn as_bool(&self) -> Option<bool> {
        match self {
            Js::Bool(b) => Some(*b),
            _ => None,
        }
    }
    /// `x === true`.
    pub fn is_true(&self) -> bool {
        matches!(self, Js::Bool(true))
    }
    pub fn as_obj(&self) -> Option<&Obj> {
        match self {
            Js::Obj(o) => Some(o),
            _ => None,
        }
    }
    pub fn as_arr(&self) -> Option<&Arr> {
        match self {
            Js::Arr(a) => Some(a),
            _ => None,
        }
    }
    pub fn as_map(&self) -> Option<&JsMap> {
        match self {
            Js::Map(m) => Some(m),
            _ => None,
        }
    }
    pub fn as_bytes(&self) -> Option<&Rc<Vec<u8>>> {
        match self {
            Js::Bytes(b) => Some(b),
            _ => None,
        }
    }

    // ----- object properties -----

    /// `x?.[key]` for an object; `undefined` for anything else or a missing key.
    pub fn get(&self, key: &str) -> Js {
        match self {
            Js::Obj(o) => o.0.borrow().get(key).cloned().unwrap_or_default(),
            Js::Arr(a) if key == "length" => Js::Num(a.0.borrow().len() as f64),
            _ => Js::Undefined,
        }
    }
    /// `x?.a?.b?.c`.
    pub fn path(&self, keys: &[&str]) -> Js {
        let mut value = self.clone();
        for key in keys {
            value = value.get(key);
        }
        value
    }
    /// `x[key] = value` on an object; ignored for anything else.
    pub fn set(&self, key: &str, value: impl Into<Js>) {
        if let Js::Obj(o) = self {
            let value = value.into();
            o.0.borrow_mut().insert(key.to_owned(), value);
        }
    }
    /// `delete x[key]`, keeping the order of the other keys.
    pub fn delete(&self, key: &str) {
        if let Js::Obj(o) = self {
            o.0.borrow_mut().shift_remove(key);
        }
    }
    /// `key in x` / `Object.hasOwn(x, key)`.
    pub fn has(&self, key: &str) -> bool {
        match self {
            Js::Obj(o) => o.0.borrow().contains_key(key),
            _ => false,
        }
    }
    /// `Object.keys(x)` for an object (array indices for an array).
    pub fn keys(&self) -> Vec<String> {
        match self {
            Js::Obj(o) => o.0.borrow().keys().cloned().collect(),
            Js::Arr(a) => (0..a.0.borrow().len()).map(|i| i.to_string()).collect(),
            _ => Vec::new(),
        }
    }
    /// `Object.entries(x)` for an object.
    pub fn entries(&self) -> Vec<(String, Js)> {
        match self {
            Js::Obj(o) => {
                o.0.borrow()
                    .iter()
                    .map(|(k, v)| (k.clone(), v.clone()))
                    .collect()
            }
            Js::Arr(a) => {
                a.0.borrow()
                    .iter()
                    .enumerate()
                    .map(|(i, v)| (i.to_string(), v.clone()))
                    .collect()
            }
            _ => Vec::new(),
        }
    }
    /// `{ ...x }`: a new object with `x`'s own properties (none for a non-object).
    pub fn spread(&self) -> Js {
        match self {
            Js::Obj(o) => Js::obj_from(o.0.borrow().clone()),
            Js::Arr(a) => Js::obj_from(
                a.0.borrow()
                    .iter()
                    .enumerate()
                    .map(|(i, v)| (i.to_string(), v.clone()))
                    .collect(),
            ),
            _ => Js::new_obj(),
        }
    }
    /// `{ ...x, ...props }`.
    pub fn with(&self, props: &[(&str, Js)]) -> Js {
        let out = self.spread();
        for (key, value) in props {
            out.set(key, value.clone());
        }
        out
    }
    /// `Object.assign(x, other)`: copies `other`'s own properties onto `x`.
    pub fn assign(&self, other: &Js) {
        for (key, value) in other.entries() {
            self.set(&key, value);
        }
    }
    /// `Object.keys(x).length`.
    pub fn key_count(&self) -> usize {
        match self {
            Js::Obj(o) => o.0.borrow().len(),
            _ => 0,
        }
    }
    pub fn props(&self) -> Option<Ref<'_, Props>> {
        match self {
            Js::Obj(o) => Some(o.0.borrow()),
            _ => None,
        }
    }
    pub fn props_mut(&self) -> Option<RefMut<'_, Props>> {
        match self {
            Js::Obj(o) => Some(o.0.borrow_mut()),
            _ => None,
        }
    }

    // ----- arrays -----

    /// `x.length` of an array or string (UTF-16 units); 0 otherwise.
    pub fn len(&self) -> usize {
        match self {
            Js::Arr(a) => a.0.borrow().len(),
            Js::Str(s) => len16(s),
            Js::Map(m) => m.0.borrow().len(),
            Js::Bytes(b) => b.len(),
            _ => 0,
        }
    }
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }
    /// `x[i]`; `undefined` past the end or for a non-array.
    pub fn at(&self, index: usize) -> Js {
        match self {
            Js::Arr(a) => a.0.borrow().get(index).cloned().unwrap_or_default(),
            _ => Js::Undefined,
        }
    }
    /// `x.at(-1)`.
    pub fn last(&self) -> Js {
        match self {
            Js::Arr(a) => a.0.borrow().last().cloned().unwrap_or_default(),
            _ => Js::Undefined,
        }
    }
    /// The array's items, a snapshot; empty for a non-array.
    pub fn items(&self) -> Vec<Js> {
        match self {
            Js::Arr(a) => a.0.borrow().clone(),
            _ => Vec::new(),
        }
    }
    /// `x.push(value)` on an array; ignored for anything else.
    pub fn push(&self, value: impl Into<Js>) {
        if let Js::Arr(a) = self {
            let value = value.into();
            a.0.borrow_mut().push(value);
        }
    }
    /// `x[i] = value` on an array, growing it with `undefined` as JS does.
    pub fn set_at(&self, index: usize, value: impl Into<Js>) {
        if let Js::Arr(a) = self {
            let value = value.into();
            let mut items = a.0.borrow_mut();
            if index >= items.len() {
                items.resize(index + 1, Js::Undefined);
            }
            items[index] = value;
        }
    }
    pub fn items_mut(&self) -> Option<RefMut<'_, Vec<Js>>> {
        match self {
            Js::Arr(a) => Some(a.0.borrow_mut()),
            _ => None,
        }
    }
    /// `[...x]`: a new array with the same items.
    pub fn copy_arr(&self) -> Js {
        Js::new_arr(self.items())
    }

    // ----- maps -----

    /// `m.get(key)` of a `Map`.
    pub fn map_get(&self, key: &str) -> Js {
        match self {
            Js::Map(m) => m.0.borrow().get(key).cloned().unwrap_or_default(),
            _ => Js::Undefined,
        }
    }
    pub fn map_has(&self, key: &str) -> bool {
        match self {
            Js::Map(m) => m.0.borrow().contains_key(key),
            _ => false,
        }
    }
    pub fn map_set(&self, key: &str, value: impl Into<Js>) {
        if let Js::Map(m) = self {
            let value = value.into();
            m.0.borrow_mut().insert(key.to_owned(), value);
        }
    }
    /// `[...m]` of a `Map`, in insertion order.
    pub fn map_entries(&self) -> Vec<(String, Js)> {
        match self {
            Js::Map(m) => {
                m.0.borrow()
                    .iter()
                    .map(|(k, v)| (k.clone(), v.clone()))
                    .collect()
            }
            _ => Vec::new(),
        }
    }
    pub fn map_keys(&self) -> Vec<String> {
        match self {
            Js::Map(m) => m.0.borrow().keys().cloned().collect(),
            _ => Vec::new(),
        }
    }

    // ----- identity and copies -----

    /// The address of an object, array, map or byte buffer, for identity tables.
    pub fn ptr(&self) -> Option<usize> {
        match self {
            Js::Obj(o) => Some(Rc::as_ptr(&o.0) as *const () as usize),
            Js::Arr(a) => Some(Rc::as_ptr(&a.0) as *const () as usize),
            Js::Map(m) => Some(Rc::as_ptr(&m.0) as *const () as usize),
            Js::Bytes(b) => Some(Rc::as_ptr(b) as *const () as usize),
            _ => None,
        }
    }
    /// `a === b`.
    pub fn strict_eq(&self, other: &Js) -> bool {
        match (self, other) {
            (Js::Undefined, Js::Undefined) | (Js::Null, Js::Null) => true,
            (Js::Bool(a), Js::Bool(b)) => a == b,
            (Js::Num(a), Js::Num(b)) => a == b,
            (Js::Str(a), Js::Str(b)) => a == b,
            (a, b) => match (a.ptr(), b.ptr()) {
                (Some(x), Some(y)) => x == y && a.type_of() == b.type_of(),
                _ => false,
            },
        }
    }
    /// `structuredClone(x)`: a deep copy, maps as maps, shared references copied once.
    pub fn deep_clone(&self) -> Js {
        fn go(value: &Js, seen: &mut HashMap<usize, Js>) -> Js {
            if let Some(p) = value.ptr()
                && let Some(copy) = seen.get(&p) {
                    return copy.clone();
                }
            match value {
                Js::Obj(o) => {
                    let out = Js::new_obj();
                    seen.insert(value.ptr().unwrap(), out.clone());
                    let entries: Vec<(String, Js)> =
                        o.0.borrow()
                            .iter()
                            .map(|(k, v)| (k.clone(), v.clone()))
                            .collect();
                    for (k, v) in entries {
                        let c = go(&v, seen);
                        out.set(&k, c);
                    }
                    out
                }
                Js::Arr(a) => {
                    let out = Js::new_arr(Vec::new());
                    seen.insert(value.ptr().unwrap(), out.clone());
                    let items = a.0.borrow().clone();
                    for v in items {
                        let c = go(&v, seen);
                        out.push(c);
                    }
                    out
                }
                Js::Map(m) => {
                    let out = Js::new_map();
                    seen.insert(value.ptr().unwrap(), out.clone());
                    let entries: Vec<(String, Js)> =
                        m.0.borrow()
                            .iter()
                            .map(|(k, v)| (k.clone(), v.clone()))
                            .collect();
                    for (k, v) in entries {
                        let c = go(&v, seen);
                        out.map_set(&k, c);
                    }
                    out
                }
                Js::Bytes(b) => Js::Bytes(Rc::new((**b).clone())),
                other => other.clone(),
            }
        }
        go(self, &mut HashMap::new())
    }

    // ----- JSON -----

    /// `JSON.stringify(x)`; `None` where JS returns `undefined`.
    pub fn stringify(&self) -> Option<String> {
        if matches!(self, Js::Undefined) {
            return None;
        }
        Some(serde_json::to_string(&Stringify(self, false)).expect("JSON of a Js value"))
    }
    /// `JSON.stringify(x, (_k, v) => v instanceof Map ? [...v] : v)`.
    pub fn stringify_maps_as_entries(&self) -> Option<String> {
        if matches!(self, Js::Undefined) {
            return None;
        }
        Some(serde_json::to_string(&Stringify(self, true)).expect("JSON of a Js value"))
    }
    /// `JSON.parse(text)`.
    pub fn parse(text: &str) -> Result<Js> {
        let mut de = serde_json::Deserializer::from_str(text);
        let value = Js::deserialize(&mut de).map_err(|e| SaveError::new(e.to_string()))?;
        de.end().map_err(|e| SaveError::new(e.to_string()))?;
        Ok(value)
    }
    /// A `serde_json::Value` of `JSON.parse(JSON.stringify(x))`.
    pub fn to_serde(&self) -> serde_json::Value {
        match self.stringify() {
            Some(text) => serde_json::from_str(&text).expect("JSON of a Js value"),
            None => serde_json::Value::Null,
        }
    }
    /// The `Js` of a `serde_json::Value`.
    pub fn from_serde(value: &serde_json::Value) -> Js {
        match value {
            serde_json::Value::Null => Js::Null,
            serde_json::Value::Bool(b) => Js::Bool(*b),
            serde_json::Value::Number(n) => Js::Num(n.as_f64().unwrap_or(f64::NAN)),
            serde_json::Value::String(s) => Js::from(s.as_str()),
            serde_json::Value::Array(items) => {
                Js::new_arr(items.iter().map(Js::from_serde).collect())
            }
            serde_json::Value::Object(map) => Js::obj_from(
                map.iter()
                    .map(|(k, v)| (k.clone(), Js::from_serde(v)))
                    .collect(),
            ),
        }
    }
}

impl fmt::Debug for Js {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Js::Undefined => f.write_str("undefined"),
            Js::Map(m) => write!(f, "Map({})", m.0.borrow().len()),
            Js::Bytes(b) => write!(f, "Bytes({})", b.len()),
            other => f.write_str(&other.stringify().unwrap_or_default()),
        }
    }
}

/// `String(n)` for a number, as ECMAScript formats it.
pub fn num_to_string(n: f64) -> String {
    if n.is_nan() {
        return "NaN".into();
    }
    if n.is_infinite() {
        return if n > 0.0 { "Infinity" } else { "-Infinity" }.into();
    }
    if n == 0.0 {
        return "0".into();
    }
    let mut buffer = ryu_js::Buffer::new();
    buffer.format_finite(n).to_owned()
}

/// `String(x)`.
pub fn to_js_string(value: &Js) -> String {
    match value {
        Js::Undefined => "undefined".into(),
        Js::Null => "null".into(),
        Js::Bool(b) => b.to_string(),
        Js::Num(n) => num_to_string(*n),
        Js::Str(s) => s.to_string(),
        Js::Arr(a) => {
            a.0.borrow()
                .iter()
                .map(|v| {
                    if v.is_nullish() {
                        String::new()
                    } else {
                        to_js_string(v)
                    }
                })
                .collect::<Vec<_>>()
                .join(",")
        }
        Js::Map(_) => "[object Map]".into(),
        _ => "[object Object]".into(),
    }
}

/// `Number(x)`.
pub fn to_number(value: &Js) -> f64 {
    match value {
        Js::Undefined => f64::NAN,
        Js::Null => 0.0,
        Js::Bool(b) => f64::from(u8::from(*b)),
        Js::Num(n) => *n,
        Js::Str(s) => {
            let t = s.trim();
            if t.is_empty() {
                0.0
            } else if let Some(hex) = t.strip_prefix("0x").or_else(|| t.strip_prefix("0X")) {
                i64::from_str_radix(hex, 16).map_or(f64::NAN, |v| v as f64)
            } else if t == "Infinity" || t == "+Infinity" {
                f64::INFINITY
            } else if t == "-Infinity" {
                f64::NEG_INFINITY
            } else if t
                .chars()
                .all(|c| c.is_ascii_digit() || matches!(c, '.' | 'e' | 'E' | '+' | '-'))
            {
                t.parse::<f64>().unwrap_or(f64::NAN)
            } else {
                f64::NAN
            }
        }
        _ => f64::NAN,
    }
}

/// `Number.isInteger(x)`.
pub fn is_integer(value: &Js) -> bool {
    matches!(value, Js::Num(n) if n.is_finite() && n.fract() == 0.0)
}

/// `Number.isFinite(x)`.
pub fn is_finite(value: &Js) -> bool {
    matches!(value, Js::Num(n) if n.is_finite())
}

/// `parseInt(s, 10)`.
pub fn parse_int(s: &str) -> f64 {
    let t = s.trim_start();
    let (sign, rest) = match t.as_bytes().first() {
        Some(b'-') => (-1.0, &t[1..]),
        Some(b'+') => (1.0, &t[1..]),
        _ => (1.0, t),
    };
    let digits: String = rest.chars().take_while(char::is_ascii_digit).collect();
    if digits.is_empty() {
        f64::NAN
    } else {
        sign * digits.parse::<f64>().unwrap_or(f64::NAN)
    }
}

/// `parseFloat(s)`.
pub fn parse_float(s: &str) -> f64 {
    let t = s.trim_start();
    let mut end = 0;
    let bytes = t.as_bytes();
    let mut seen_digit = false;
    let mut seen_dot = false;
    let mut seen_exp = false;
    while end < bytes.len() {
        let c = bytes[end];
        if c.is_ascii_digit() {
            seen_digit = true;
        } else if c == b'.' && !seen_dot && !seen_exp {
            seen_dot = true;
        } else if (c == b'e' || c == b'E') && seen_digit && !seen_exp {
            let next = bytes.get(end + 1).copied();
            let after = if matches!(next, Some(b'+' | b'-')) {
                bytes.get(end + 2).copied()
            } else {
                next
            };
            if !after.is_some_and(|b| b.is_ascii_digit()) {
                break;
            }
            seen_exp = true;
            if matches!(next, Some(b'+' | b'-')) {
                end += 1;
            }
        } else if (c == b'+' || c == b'-') && end == 0 {
        } else {
            break;
        }
        end += 1;
    }
    if t.starts_with("Infinity") || t.starts_with("+Infinity") {
        return f64::INFINITY;
    }
    if t.starts_with("-Infinity") {
        return f64::NEG_INFINITY;
    }
    t[..end].parse::<f64>().unwrap_or(f64::NAN)
}

// ----- UTF-16 strings -----

/// `s.length`: UTF-16 code units.
pub fn len16(s: &str) -> usize {
    s.chars().map(char::len_utf16).sum()
}

/// The UTF-16 code units of `s`.
pub fn utf16(s: &str) -> Vec<u16> {
    s.encode_utf16().collect()
}

/// A string from UTF-16 code units; a lone surrogate becomes U+FFFD.
pub fn from_utf16(units: &[u16]) -> String {
    String::from_utf16_lossy(units)
}

fn rel_index(index: f64, len: usize) -> usize {
    if index.is_nan() {
        0
    } else if index < 0.0 {
        (len as f64 + index).max(0.0) as usize
    } else {
        (index as usize).min(len)
    }
}

/// `s.slice(start, end)` in UTF-16 units; `end` `None` for the rest.
pub fn slice16(s: &str, start: f64, end: Option<f64>) -> String {
    let units = utf16(s);
    let len = units.len();
    let from = rel_index(start, len);
    let to = end.map_or(len, |e| rel_index(e, len));
    if from >= to {
        String::new()
    } else {
        from_utf16(&units[from..to])
    }
}

/// `s.charCodeAt(i)`, `None` past the end.
pub fn char_code_at(s: &str, index: usize) -> Option<u16> {
    s.encode_utf16().nth(index)
}

/// `s.indexOf(needle, from)` in UTF-16 units; `None` for -1.
pub fn index_of16(s: &str, needle: &str, from: usize) -> Option<usize> {
    let hay = utf16(s);
    let pin = utf16(needle);
    if pin.is_empty() {
        return Some(from.min(hay.len()));
    }
    if pin.len() > hay.len() {
        return None;
    }
    (from..=hay.len() - pin.len()).find(|&i| hay[i..i + pin.len()] == pin[..])
}

/// `a < b` for strings: UTF-16 code unit order.
pub fn cmp16(a: &str, b: &str) -> std::cmp::Ordering {
    a.encode_utf16().cmp(b.encode_utf16())
}

// ----- identity tables (WeakMap / WeakSet) -----

enum WeakRef {
    Obj(Weak<RefCell<Props>>),
    Arr(Weak<RefCell<Vec<Js>>>),
    Map(Weak<RefCell<StrMap>>),
    Bytes(Weak<Vec<u8>>),
}

impl WeakRef {
    fn of(value: &Js) -> Option<Self> {
        Some(match value {
            Js::Obj(o) => WeakRef::Obj(Rc::downgrade(&o.0)),
            Js::Arr(a) => WeakRef::Arr(Rc::downgrade(&a.0)),
            Js::Map(m) => WeakRef::Map(Rc::downgrade(&m.0)),
            Js::Bytes(b) => WeakRef::Bytes(Rc::downgrade(b)),
            _ => return None,
        })
    }
    fn alive(&self) -> bool {
        match self {
            WeakRef::Obj(w) => w.strong_count() > 0,
            WeakRef::Arr(w) => w.strong_count() > 0,
            WeakRef::Map(w) => w.strong_count() > 0,
            WeakRef::Bytes(w) => w.strong_count() > 0,
        }
    }
}

/// A `WeakMap` keyed by object identity (a `WeakSet` with `V = ()`).
pub struct Marks<V> {
    entries: RefCell<HashMap<usize, (WeakRef, V)>>,
    /// The size at which dead entries are next dropped: twice the live ones at the last prune.
    prune_at: std::cell::Cell<usize>,
}

impl<V: Clone> Default for Marks<V> {
    fn default() -> Self {
        Self::new()
    }
}

impl<V: Clone> Marks<V> {
    pub fn new() -> Self {
        Self {
            entries: RefCell::new(HashMap::new()),
            prune_at: std::cell::Cell::new(1024),
        }
    }
    pub fn set(&self, key: &Js, value: V) {
        if let (Some(p), Some(w)) = (key.ptr(), WeakRef::of(key)) {
            let dead = {
                let mut entries = self.entries.borrow_mut();
                let dead = if entries.len() >= self.prune_at.get() {
                    let (live, dead): (HashMap<_, _>, HashMap<_, _>) =
                        entries.drain().partition(|(_, (w, _))| w.alive());
                    *entries = live;
                    self.prune_at.set((entries.len() * 2).max(1024));
                    dead
                } else {
                    HashMap::new()
                };
                entries.insert(p, (w, value));
                dead
            };
            drop(dead);
        }
    }
    /// Drops the entries of objects that no longer exist, and their values.
    pub fn prune(&self) {
        let dead: HashMap<_, _> = {
            let mut entries = self.entries.borrow_mut();
            let (live, dead) = entries.drain().partition(|(_, (w, _))| w.alive());
            *entries = live;
            dead
        };
        drop(dead);
    }
    pub fn get(&self, key: &Js) -> Option<V> {
        let p = key.ptr()?;
        let entries = self.entries.borrow();
        let (w, v) = entries.get(&p)?;
        w.alive().then(|| v.clone())
    }
    pub fn has(&self, key: &Js) -> bool {
        self.get(key).is_some()
    }
    pub fn delete(&self, key: &Js) {
        if let Some(p) = key.ptr() {
            self.entries.borrow_mut().remove(&p);
        }
    }
    pub fn clear(&self) {
        self.entries.borrow_mut().clear();
    }
}

// ----- serde -----

struct Stringify<'a>(&'a Js, bool);

impl Serialize for Stringify<'_> {
    fn serialize<S: Serializer>(&self, s: S) -> std::result::Result<S::Ok, S::Error> {
        let maps_as_entries = self.1;
        match self.0 {
            Js::Undefined | Js::Null => s.serialize_unit(),
            Js::Bool(b) => s.serialize_bool(*b),
            Js::Num(n) => {
                if !n.is_finite() {
                    s.serialize_unit()
                } else if n.fract() == 0.0 && n.abs() < 9.007_199_254_740_992e15 {
                    s.serialize_i64(*n as i64)
                } else {
                    s.serialize_f64(*n)
                }
            }
            Js::Str(v) => s.serialize_str(v),
            Js::Arr(a) => {
                let items = a.0.borrow();
                let mut seq = s.serialize_seq(Some(items.len()))?;
                for item in items.iter() {
                    seq.serialize_element(&Stringify(item, maps_as_entries))?;
                }
                seq.end()
            }
            Js::Obj(o) => {
                let props = o.0.borrow();
                let mut map = s.serialize_map(None)?;
                for (k, v) in props.iter() {
                    if !v.is_undefined() {
                        map.serialize_entry(k, &Stringify(v, maps_as_entries))?;
                    }
                }
                map.end()
            }
            Js::Map(m) if maps_as_entries => {
                let entries = m.0.borrow();
                let mut seq = s.serialize_seq(Some(entries.len()))?;
                for (k, v) in entries.iter() {
                    let pair = Js::new_arr(vec![Js::from(k.as_str()), v.clone()]);
                    seq.serialize_element(&Stringify(&pair, maps_as_entries))?;
                }
                seq.end()
            }
            Js::Map(_) | Js::Bytes(_) => s.serialize_map(Some(0))?.end(),
        }
    }
}

impl Serialize for Js {
    fn serialize<S: Serializer>(&self, s: S) -> std::result::Result<S::Ok, S::Error> {
        Stringify(self, false).serialize(s)
    }
}

impl<'de> Deserialize<'de> for Js {
    fn deserialize<D: Deserializer<'de>>(d: D) -> std::result::Result<Js, D::Error> {
        struct V;
        impl<'de> Visitor<'de> for V {
            type Value = Js;
            fn expecting(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
                f.write_str("a JSON value")
            }
            fn visit_bool<E: de::Error>(self, v: bool) -> std::result::Result<Js, E> {
                Ok(Js::Bool(v))
            }
            fn visit_i64<E: de::Error>(self, v: i64) -> std::result::Result<Js, E> {
                Ok(Js::Num(v as f64))
            }
            fn visit_u64<E: de::Error>(self, v: u64) -> std::result::Result<Js, E> {
                Ok(Js::Num(v as f64))
            }
            fn visit_f64<E: de::Error>(self, v: f64) -> std::result::Result<Js, E> {
                Ok(Js::Num(v))
            }
            fn visit_str<E: de::Error>(self, v: &str) -> std::result::Result<Js, E> {
                Ok(Js::from(v))
            }
            fn visit_string<E: de::Error>(self, v: String) -> std::result::Result<Js, E> {
                Ok(Js::from(v))
            }
            fn visit_unit<E: de::Error>(self) -> std::result::Result<Js, E> {
                Ok(Js::Null)
            }
            fn visit_none<E: de::Error>(self) -> std::result::Result<Js, E> {
                Ok(Js::Null)
            }
            fn visit_some<D: Deserializer<'de>>(self, d: D) -> std::result::Result<Js, D::Error> {
                Js::deserialize(d)
            }
            fn visit_seq<A: SeqAccess<'de>>(self, mut seq: A) -> std::result::Result<Js, A::Error> {
                let mut items = Vec::with_capacity(seq.size_hint().unwrap_or(0));
                while let Some(item) = seq.next_element::<Js>()? {
                    items.push(item);
                }
                Ok(Js::new_arr(items))
            }
            fn visit_map<A: MapAccess<'de>>(self, mut map: A) -> std::result::Result<Js, A::Error> {
                let mut props = Props::with_capacity_and_hasher(
                    map.size_hint().unwrap_or(0),
                    Hasher::default(),
                );
                while let Some((k, v)) = map.next_entry::<String, Js>()? {
                    props.insert(k, v);
                }
                Ok(Js::obj_from(props))
            }
        }
        d.deserialize_any(V)
    }
}

/// Runs `f`, printing how long it took under `label` when `NATIVE_SAVE_TRACE` is set.
pub fn timed<T>(label: &str, f: impl FnOnce() -> T) -> T {
    if std::env::var_os("NATIVE_SAVE_TRACE").is_none() {
        return f();
    }
    let t = std::time::Instant::now();
    let out = f();
    eprintln!("  {label}: {:.1} ms", t.elapsed().as_secs_f64() * 1e3);
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn stringify_follows_js() {
        let o = obj! { "b" => 1.0, "a" => Js::Undefined, "c" => arr![Js::Undefined, 1.5, "x"] };
        assert_eq!(o.stringify().unwrap(), r#"{"b":1,"c":[null,1.5,"x"]}"#);
        let parsed = Js::parse(r#"{"z":1,"a":{"y":[true,null]}}"#).unwrap();
        assert_eq!(
            parsed.stringify().unwrap(),
            r#"{"z":1,"a":{"y":[true,null]}}"#
        );
        assert_eq!(num_to_string(1e21), "1e+21");
        assert_eq!(num_to_string(0.1 + 0.2), "0.30000000000000004");
        assert_eq!(num_to_string(-0.0), "0");
    }

    #[test]
    fn identity_and_spread() {
        let a = obj! { "x" => 1.0 };
        let b = a.clone();
        assert!(a.strict_eq(&b));
        let c = a.spread();
        assert!(!a.strict_eq(&c));
        c.set("x", 2.0);
        assert_eq!(a.get("x").as_f64(), Some(1.0));
        b.set("x", 3.0);
        assert_eq!(a.get("x").as_f64(), Some(3.0));
        let marks = Marks::<u32>::new();
        marks.set(&a, 7);
        assert_eq!(marks.get(&b), Some(7));
        assert_eq!(marks.get(&c), None);
    }

    #[test]
    fn utf16_strings() {
        let s = "a😀b";
        assert_eq!(len16(s), 4);
        assert_eq!(slice16(s, 1.0, Some(3.0)), "😀");
        assert_eq!(slice16(s, -1.0, None), "b");
        assert_eq!(index_of16(s, "b", 0), Some(3));
    }
}
