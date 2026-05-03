# Apache POI Equation3 Native-Stream Fixtures

This directory contains minimal Apache-derived Equation Editor 3.0 `Equation Native` stream controls.
It does not vendor the full source Word `.doc` files.

Source project:

- Apache POI
- Source commit: `e6a04b49211e23c704fcdbe524d99d2f4486b083`
- License: Apache-2.0
- License and notice retention:
  - `THIRD_PARTY_LICENSES/Apache-POI-LICENSE.txt`
  - `THIRD_PARTY_NOTICES/Apache-POI-NOTICE.txt`

Included controls:

- `apache_poi_bug61268_formula0001.equation-native.hex`
  - Converts under the implemented limited MTEF v3 path.
- `apache_poi_bug61268_formula0003.equation-native.hex`
  - Converts under the implemented limited MTEF v3 path and exercises an observed bracket template.
- `apache_poi_bug50936_1_formula0013_selector43v2.equation-native.hex`
  - Preserves an observed unsupported `selector=43 variation=2` control case.

`SOURCES.json` records the fixed upstream URLs, source document SHA-256 values, OLE stream names, native stream hashes, expected conversion status, and canonical hashes for the converting controls.

These fixtures are static native streams. The tests do not open Word, execute macros, instantiate an OLE server, or use external converter tools.
