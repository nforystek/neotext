xof 0303txt 0032

// Generated by 3D Rad Exporter plugin for Google SketchUp - http://www.3drad.com

template Header {
<3D82AB43-62DA-11cf-AB39-0020AF71E433>
WORD major;
WORD minor;
DWORD flags;
}
template Vector {
<3D82AB5E-62DA-11cf-AB39-0020AF71E433>
FLOAT x;
FLOAT y;
FLOAT z;
}
template Coords2d {
<F6F23F44-7686-11cf-8F52-0040333594A3>
FLOAT u;
FLOAT v;
}
template Matrix4x4 {
<F6F23F45-7686-11cf-8F52-0040333594A3>
array FLOAT matrix[16];
}
template ColorRGBA {
<35FF44E0-6C7C-11cf-8F52-0040333594A3>
FLOAT red;
FLOAT green;
FLOAT blue;
FLOAT alpha;
}
template ColorRGB {
<D3E16E81-7835-11cf-8F52-0040333594A3>
FLOAT red;
FLOAT green;
FLOAT blue;
}
template IndexedColor {
<1630B820-7842-11cf-8F52-0040333594A3>
DWORD index;
ColorRGBA indexColor;
}
template Boolean {
<4885AE61-78E8-11cf-8F52-0040333594A3>
WORD truefalse;
}
template Boolean2d {
<4885AE63-78E8-11cf-8F52-0040333594A3>
Boolean u;
Boolean v;
}
template MaterialWrap {
<4885AE60-78E8-11cf-8F52-0040333594A3>
Boolean u;
Boolean v;
}
template TextureFilename {
<A42790E1-7810-11cf-8F52-0040333594A3>
STRING filename;
}
template Material {
<3D82AB4D-62DA-11cf-AB39-0020AF71E433>
ColorRGBA faceColor;
FLOAT power;
ColorRGB specularColor;
ColorRGB emissiveColor;
[...]
}
template MeshFace {
<3D82AB5F-62DA-11cf-AB39-0020AF71E433>
DWORD nFaceVertexIndices;
array DWORD faceVertexIndices[nFaceVertexIndices];
}
template MeshFaceWraps {
<4885AE62-78E8-11cf-8F52-0040333594A3>
DWORD nFaceWrapValues;
Boolean2d faceWrapValues;
}
template MeshTextureCoords {
<F6F23F40-7686-11cf-8F52-0040333594A3>
DWORD nTextureCoords;
array Coords2d textureCoords[nTextureCoords];
}
template MeshMaterialList {
<F6F23F42-7686-11cf-8F52-0040333594A3>
DWORD nMaterials;
DWORD nFaceIndexes;
array DWORD faceIndexes[nFaceIndexes];
[Material]
}
template MeshNormals {
<F6F23F43-7686-11cf-8F52-0040333594A3>
DWORD nNormals;
array Vector normals[nNormals];
DWORD nFaceNormals;
array MeshFace faceNormals[nFaceNormals];
}
template MeshVertexColors {
<1630B821-7842-11cf-8F52-0040333594A3>
DWORD nVertexColors;
array IndexedColor vertexColors[nVertexColors];
}
template Mesh {
<3D82AB44-62DA-11cf-AB39-0020AF71E433>
DWORD nVertices;
array Vector vertices[nVertices];
DWORD nFaces;
array MeshFace faces[nFaces];
[...]
}
template FrameTransformMatrix {
<F6F23F41-7686-11cf-8F52-0040333594A3>
Matrix4x4 frameMatrix;
}
template Frame {
<3D82AB46-62DA-11cf-AB39-0020AF71E433>
[...]
}
template XSkinMeshHeader {
<3cf169ce-ff7c-44ab-93c0-f78f62d172e2>
WORD nMaxSkinWeightsPerVertex;
WORD nMaxSkinWeightsPerFace;
WORD nBones;
}
template VertexDuplicationIndices {
<b8d65549-d7c9-4995-89cf-53a9a8b031e3>
DWORD nIndices;
DWORD nOriginalVertices;
array DWORD indices[nIndices];
}
template SkinWeights {
<6f0d123b-bad2-4167-a0d0-80224f25fabb>
STRING transformNodeName;
DWORD nWeights;
array DWORD vertexIndices[nWeights];
array FLOAT weights[nWeights];
Matrix4x4 matrixOffset;
}
Frame RAD_SCENE_ROOT {
FrameTransformMatrix {
1.000000,0.000000,0.000000,0.000000,0.000000,1.000000,0.000000,0.000000,0.000000,0.000000,1.000000,0.000000,0.000000,0.000000,0.000000,1.000000;;
}
Frame RAD_FRAME {
FrameTransformMatrix {
1.000000,0.000000,0.000000,0.000000,0.000000,1.000000,0.000000,0.000000,0.000000,0.000000,1.000000,0.000000,0.000000,0.000000,0.000000,1.000000;;
}
Mesh RAD_MESH {
295;
4.819690;0.250210;3.386844;,
4.452649;0.250210;3.669903;,
4.452649;0.250210;3.386844;,
4.452649;0.250210;3.819464;,
4.819690;0.250210;3.819464;,
4.819690;0.250210;3.669903;,
4.819690;-2.705874;3.386844;,
4.452649;-2.705874;3.669903;,
4.452649;-2.705874;3.386844;,
4.452649;-2.705874;3.819464;,
4.819690;-2.705874;3.819464;,
4.819690;-2.705874;3.519690;,
4.819690;-2.705874;3.669903;,
-0.852159;0.250210;3.386844;,
-1.219199;0.250210;3.669903;,
-1.219199;0.250210;3.386844;,
-1.219199;0.250210;3.819464;,
-0.852159;0.250210;3.819464;,
-0.852159;0.250210;3.669903;,
-0.852159;-0.936071;3.386844;,
-1.219199;-0.936071;3.669903;,
-1.219199;-0.936071;3.386844;,
-1.219199;-0.936071;3.819464;,
-0.852159;-0.936071;3.819464;,
-0.852159;-0.936071;3.669903;,
-0.852159;-0.342930;3.386844;,
-1.219199;-0.342930;3.669903;,
-1.219199;-0.342930;3.386844;,
-1.219199;-0.342930;3.819464;,
-0.852159;-0.342930;3.819464;,
-0.852159;-0.342930;3.669903;,
-0.852159;-0.046360;3.386844;,
-1.219199;-0.046360;3.669903;,
-1.219199;-0.046360;3.386844;,
-1.219199;-0.046360;3.819464;,
-0.852159;-0.046360;3.819464;,
-0.852159;-0.046360;3.669903;,
-0.852159;-0.639501;3.386844;,
-1.219199;-0.639501;3.669903;,
-1.219199;-0.639501;3.386844;,
-1.219199;-0.639501;3.819464;,
-0.852159;-0.639501;3.819464;,
-0.852159;-0.639501;3.669903;,
4.819690;-1.227832;3.386844;,
4.452649;-1.227832;3.669903;,
4.452649;-1.227832;3.386844;,
4.452649;-1.227832;3.819464;,
4.819690;-1.227832;3.819464;,
4.819690;-1.227832;3.669903;,
4.819690;-0.488811;3.386844;,
4.452649;-0.488811;3.669903;,
4.452649;-0.488811;3.386844;,
4.452649;-0.488811;3.819464;,
4.819690;-0.488811;3.819464;,
4.819690;-0.488811;3.669903;,
4.819690;-1.966853;3.386844;,
4.452649;-1.966853;3.669903;,
4.452649;-1.966853;3.386844;,
4.452649;-1.966853;3.819464;,
4.819690;-1.966853;3.819464;,
4.819690;-1.966853;3.669903;,
4.819690;-1.597343;3.386844;,
4.452649;-1.597343;3.669903;,
4.452649;-1.597343;3.386844;,
4.452649;-1.597343;3.819464;,
4.819690;-1.597343;3.819464;,
4.819690;-1.597343;3.669903;,
4.819690;-2.336364;3.386844;,
4.452649;-2.336364;3.669903;,
4.452649;-2.336364;3.386844;,
4.452649;-2.336364;3.819464;,
4.819690;-2.336364;3.819464;,
4.819690;-2.336364;3.669903;,
4.819690;-0.858322;3.386844;,
4.452649;-0.858322;3.669903;,
4.452649;-0.858322;3.386844;,
4.452649;-0.858322;3.819464;,
4.819690;-0.858322;3.819464;,
4.819690;-0.858322;3.669903;,
4.819690;-0.119301;3.386844;,
4.452649;-0.119301;3.669903;,
4.452649;-0.119301;3.386844;,
4.452649;-0.119301;3.819464;,
4.819690;-0.119301;3.819464;,
4.819690;-0.119301;3.669903;,
-0.852159;0.546780;3.386844;,
-1.219199;0.546780;3.669903;,
-1.219199;0.546780;3.386844;,
-1.219199;0.546780;4.063692;,
-0.852159;0.546780;4.063692;,
-0.852159;0.546780;3.669903;,
4.819690;0.619720;3.386844;,
4.452649;0.619720;3.669903;,
4.452649;0.619720;3.386844;,
4.452649;0.619720;4.063692;,
4.819690;0.619720;4.063692;,
4.819690;0.619720;3.669903;,
-43.080017;16.281542;19.674905;,
-43.522865;16.281542;21.282653;,
-49.879648;16.281542;17.801974;,
-43.522865;16.281542;21.282653;,
-50.255143;16.281542;19.165201;,
-49.879648;16.281542;17.801974;,
-43.080017;-3.308203;19.674905;,
-43.522865;-3.308203;21.282653;,
-49.879648;-3.308203;17.801974;,
-43.522865;-3.308203;21.282653;,
-50.255143;-3.308203;19.165201;,
-49.879648;-3.308203;17.801974;,
-43.080017;6.486670;19.674905;,
-43.522865;6.486670;21.282653;,
-49.879648;6.486670;17.801974;,
-43.522865;6.486670;21.282653;,
-50.255143;6.486670;19.165201;,
-49.879648;6.486670;17.801974;,
-43.080017;11.384106;19.674905;,
-43.522865;11.384106;21.282653;,
-49.879648;11.384106;17.801974;,
-43.522865;11.384106;21.282653;,
-50.255143;11.384106;19.165201;,
-49.879648;11.384106;17.801974;,
-43.080017;1.589233;19.674905;,
-43.522865;1.589233;21.282653;,
-49.879648;1.589233;17.801974;,
-43.522865;1.589233;21.282653;,
-50.255143;1.589233;19.165201;,
-49.879648;1.589233;17.801974;,
-43.080017;13.832824;19.674905;,
-43.522865;13.832824;21.282653;,
-49.879648;13.832824;17.801974;,
-43.522865;13.832824;21.282653;,
-50.255143;13.832824;19.165201;,
-49.879648;13.832824;17.801974;,
-43.080017;8.935388;19.674905;,
-43.522865;8.935388;21.282653;,
-49.879648;8.935388;17.801974;,
-43.522865;8.935388;21.282653;,
-50.255143;8.935388;19.165201;,
-49.879648;8.935388;17.801974;,
-43.522865;4.037951;21.282653;,
-50.255143;4.037951;19.165201;,
-49.879648;4.037951;17.801974;,
-43.080017;4.037951;19.674905;,
-43.522865;4.037951;21.282653;,
-49.879648;4.037951;17.801974;,
-43.080017;-0.859485;19.674905;,
-43.522865;-0.859485;21.282653;,
-49.879648;-0.859485;17.801974;,
-43.522865;-0.859485;21.282653;,
-50.255143;-0.859485;19.165201;,
-49.879648;-0.859485;17.801974;,
-43.080017;15.057183;19.674905;,
-43.522865;15.057183;21.282653;,
-49.879648;15.057183;17.801974;,
-43.522865;15.057183;21.282653;,
-50.255143;15.057183;19.165201;,
-49.879648;15.057183;17.801974;,
-43.080017;12.608465;19.674905;,
-43.522865;12.608465;21.282653;,
-49.879648;12.608465;17.801974;,
-43.522865;12.608465;21.282653;,
-50.255143;12.608465;19.165201;,
-49.879648;12.608465;17.801974;,
-43.080017;10.159747;19.674905;,
-43.522865;10.159747;21.282653;,
-49.879648;10.159747;17.801974;,
-43.522865;10.159747;21.282653;,
-50.255143;10.159747;19.165201;,
-49.879648;10.159747;17.801974;,
-43.522865;7.711029;21.282653;,
-50.255143;7.711029;19.165201;,
-49.879648;7.711029;17.801974;,
-43.080017;7.711029;19.674905;,
-43.522865;7.711029;21.282653;,
-49.879648;7.711029;17.801974;,
-43.522865;5.262310;21.282653;,
-50.255143;5.262310;19.165201;,
-49.879648;5.262310;17.801974;,
-43.080017;5.262310;19.674905;,
-43.522865;5.262310;21.282653;,
-49.879648;5.262310;17.801974;,
-43.080017;2.813592;19.674905;,
-43.522865;2.813592;21.282653;,
-49.879648;2.813592;17.801974;,
-43.522865;2.813592;21.282653;,
-50.255143;2.813592;19.165201;,
-49.879648;2.813592;17.801974;,
-43.080017;0.364874;19.674905;,
-43.522865;0.364874;21.282653;,
-49.879648;0.364874;17.801974;,
-43.522865;0.364874;21.282653;,
-50.255143;0.364874;19.165201;,
-49.879648;0.364874;17.801974;,
-43.080017;-2.083844;19.674905;,
-43.522865;-2.083844;21.282653;,
-49.879648;-2.083844;17.801974;,
-43.522865;-2.083844;21.282653;,
-50.255143;-2.083844;19.165201;,
-49.879648;-2.083844;17.801974;,
-43.080017;-2.696023;19.674905;,
-43.522865;-2.696023;21.282653;,
-49.879648;-2.696023;17.801974;,
-43.522865;-2.696023;21.282653;,
-50.255143;-2.696023;19.165201;,
-49.879648;-2.696023;17.801974;,
-43.080017;-1.471664;19.674905;,
-43.522865;-1.471664;21.282653;,
-49.879648;-1.471664;17.801974;,
-43.522865;-1.471664;21.282653;,
-50.255143;-1.471664;19.165201;,
-49.879648;-1.471664;17.801974;,
-43.080017;-0.247305;19.674905;,
-43.522865;-0.247305;21.282653;,
-49.879648;-0.247305;17.801974;,
-43.522865;-0.247305;21.282653;,
-50.255143;-0.247305;19.165201;,
-49.879648;-0.247305;17.801974;,
-43.080017;0.977054;19.674905;,
-43.522865;0.977054;21.282653;,
-49.879648;0.977054;17.801974;,
-43.522865;0.977054;21.282653;,
-50.255143;0.977054;19.165201;,
-49.879648;0.977054;17.801974;,
-43.522865;2.201413;21.282653;,
-50.255143;2.201413;19.165201;,
-49.879648;2.201413;17.801974;,
-43.080017;2.201413;19.674905;,
-43.522865;2.201413;21.282653;,
-49.879648;2.201413;17.801974;,
-43.080017;3.425772;19.674905;,
-43.522865;3.425772;21.282653;,
-49.879648;3.425772;17.801974;,
-43.522865;3.425772;21.282653;,
-50.255143;3.425772;19.165201;,
-49.879648;3.425772;17.801974;,
-43.080017;4.650131;19.674905;,
-43.522865;4.650131;21.282653;,
-49.879648;4.650131;17.801974;,
-43.522865;4.650131;21.282653;,
-50.255143;4.650131;19.165201;,
-49.879648;4.650131;17.801974;,
-43.080017;5.874490;19.674905;,
-43.522865;5.874490;21.282653;,
-49.879648;5.874490;17.801974;,
-43.522865;5.874490;21.282653;,
-50.255143;5.874490;19.165201;,
-49.879648;5.874490;17.801974;,
-43.080017;7.098849;19.674905;,
-43.522865;7.098849;21.282653;,
-49.879648;7.098849;17.801974;,
-43.522865;7.098849;21.282653;,
-50.255143;7.098849;19.165201;,
-49.879648;7.098849;17.801974;,
-43.080017;8.323208;19.674905;,
-43.522865;8.323208;21.282653;,
-49.879648;8.323208;17.801974;,
-43.522865;8.323208;21.282653;,
-50.255143;8.323208;19.165201;,
-49.879648;8.323208;17.801974;,
-43.080017;9.547567;19.674905;,
-43.522865;9.547567;21.282653;,
-49.879648;9.547567;17.801974;,
-43.522865;9.547567;21.282653;,
-50.255143;9.547567;19.165201;,
-49.879648;9.547567;17.801974;,
-43.080017;10.771926;19.674905;,
-43.522865;10.771926;21.282653;,
-49.879648;10.771926;17.801974;,
-43.522865;10.771926;21.282653;,
-50.255143;10.771926;19.165201;,
-49.879648;10.771926;17.801974;,
-43.080017;11.996285;19.674905;,
-43.522865;11.996285;21.282653;,
-49.879648;11.996285;17.801974;,
-43.522865;11.996285;21.282653;,
-50.255143;11.996285;19.165201;,
-49.879648;11.996285;17.801974;,
-43.080017;13.220644;19.674905;,
-43.522865;13.220644;21.282653;,
-49.879648;13.220644;17.801974;,
-43.522865;13.220644;21.282653;,
-50.255143;13.220644;19.165201;,
-49.879648;13.220644;17.801974;,
-43.080017;14.445003;19.674905;,
-43.522865;14.445003;21.282653;,
-49.879648;14.445003;17.801974;,
-43.522865;14.445003;21.282653;,
-50.255143;14.445003;19.165201;,
-49.879648;14.445003;17.801974;,
-43.080017;15.669362;19.674905;,
-43.522865;15.669362;21.282653;,
-49.879648;15.669362;17.801974;,
-43.522865;15.669362;21.282653;,
-50.255143;15.669362;19.165201;,
-49.879648;15.669362;17.801974;;
131;
3;2,1,0,
3;3,0,1,
3;4,0,3,
3;5,0,4,
3;8,7,6,
3;9,6,7,
3;10,6,9,
3;11,6,10,
3;12,11,10,
3;15,14,13,
3;16,13,14,
3;17,13,16,
3;18,13,17,
3;21,20,19,
3;22,19,20,
3;23,19,22,
3;24,19,23,
3;27,26,25,
3;28,25,26,
3;29,25,28,
3;30,25,29,
3;33,32,31,
3;34,31,32,
3;35,31,34,
3;36,31,35,
3;39,38,37,
3;40,37,38,
3;41,37,40,
3;42,37,41,
3;45,44,43,
3;46,43,44,
3;47,43,46,
3;48,43,47,
3;51,50,49,
3;52,49,50,
3;53,49,52,
3;54,49,53,
3;57,56,55,
3;58,55,56,
3;59,55,58,
3;60,55,59,
3;63,62,61,
3;64,61,62,
3;65,61,64,
3;66,61,65,
3;69,68,67,
3;70,67,68,
3;71,67,70,
3;72,67,71,
3;75,74,73,
3;76,73,74,
3;77,73,76,
3;78,73,77,
3;81,80,79,
3;82,79,80,
3;83,79,82,
3;84,79,83,
3;87,86,85,
3;88,85,86,
3;89,85,88,
3;90,85,89,
3;93,92,91,
3;94,91,92,
3;95,91,94,
3;96,91,95,
3;99,98,97,
3;102,101,100,
3;105,104,103,
3;108,107,106,
3;111,110,109,
3;114,113,112,
3;117,116,115,
3;120,119,118,
3;123,122,121,
3;126,125,124,
3;129,128,127,
3;132,131,130,
3;135,134,133,
3;138,137,136,
3;141,140,139,
3;144,143,142,
3;147,146,145,
3;150,149,148,
3;153,152,151,
3;156,155,154,
3;159,158,157,
3;162,161,160,
3;165,164,163,
3;168,167,166,
3;171,170,169,
3;174,173,172,
3;177,176,175,
3;180,179,178,
3;183,182,181,
3;186,185,184,
3;189,188,187,
3;192,191,190,
3;195,194,193,
3;198,197,196,
3;201,200,199,
3;204,203,202,
3;207,206,205,
3;210,209,208,
3;213,212,211,
3;216,215,214,
3;219,218,217,
3;222,221,220,
3;225,224,223,
3;228,227,226,
3;231,230,229,
3;234,233,232,
3;237,236,235,
3;240,239,238,
3;243,242,241,
3;246,245,244,
3;249,248,247,
3;252,251,250,
3;255,254,253,
3;258,257,256,
3;261,260,259,
3;264,263,262,
3;267,266,265,
3;270,269,268,
3;273,272,271,
3;276,275,274,
3;279,278,277,
3;282,281,280,
3;285,284,283,
3;288,287,286,
3;291,290,289,
3;294,293,292;;
MeshNormals {
295;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;
0.000000;1.000000;0.000000;;
131;
3;2,1,0;
3;3,0,1;
3;4,0,3;
3;5,0,4;
3;8,7,6;
3;9,6,7;
3;10,6,9;
3;11,6,10;
3;12,11,10;
3;15,14,13;
3;16,13,14;
3;17,13,16;
3;18,13,17;
3;21,20,19;
3;22,19,20;
3;23,19,22;
3;24,19,23;
3;27,26,25;
3;28,25,26;
3;29,25,28;
3;30,25,29;
3;33,32,31;
3;34,31,32;
3;35,31,34;
3;36,31,35;
3;39,38,37;
3;40,37,38;
3;41,37,40;
3;42,37,41;
3;45,44,43;
3;46,43,44;
3;47,43,46;
3;48,43,47;
3;51,50,49;
3;52,49,50;
3;53,49,52;
3;54,49,53;
3;57,56,55;
3;58,55,56;
3;59,55,58;
3;60,55,59;
3;63,62,61;
3;64,61,62;
3;65,61,64;
3;66,61,65;
3;69,68,67;
3;70,67,68;
3;71,67,70;
3;72,67,71;
3;75,74,73;
3;76,73,74;
3;77,73,76;
3;78,73,77;
3;81,80,79;
3;82,79,80;
3;83,79,82;
3;84,79,83;
3;87,86,85;
3;88,85,86;
3;89,85,88;
3;90,85,89;
3;93,92,91;
3;94,91,92;
3;95,91,94;
3;96,91,95;
3;99,98,97;
3;102,101,100;
3;105,104,103;
3;108,107,106;
3;111,110,109;
3;114,113,112;
3;117,116,115;
3;120,119,118;
3;123,122,121;
3;126,125,124;
3;129,128,127;
3;132,131,130;
3;135,134,133;
3;138,137,136;
3;141,140,139;
3;144,143,142;
3;147,146,145;
3;150,149,148;
3;153,152,151;
3;156,155,154;
3;159,158,157;
3;162,161,160;
3;165,164,163;
3;168,167,166;
3;171,170,169;
3;174,173,172;
3;177,176,175;
3;180,179,178;
3;183,182,181;
3;186,185,184;
3;189,188,187;
3;192,191,190;
3;195,194,193;
3;198,197,196;
3;201,200,199;
3;204,203,202;
3;207,206,205;
3;210,209,208;
3;213,212,211;
3;216,215,214;
3;219,218,217;
3;222,221,220;
3;225,224,223;
3;228,227,226;
3;231,230,229;
3;234,233,232;
3;237,236,235;
3;240,239,238;
3;243,242,241;
3;246,245,244;
3;249,248,247;
3;252,251,250;
3;255,254,253;
3;258,257,256;
3;261,260,259;
3;264,263,262;
3;267,266,265;
3;270,269,268;
3;273,272,271;
3;276,275,274;
3;279,278,277;
3;282,281,280;
3;285,284,283;
3;288,287,286;
3;291,290,289;
3;294,293,292;;
}
MeshTextureCoords {
295;
190.751679,-133.340374;
176.301252,-144.484443;
176.301252,-133.340374;
176.301252,-150.372674;
190.751679,-150.372674;
190.751679,-144.484443;
190.751679,-133.340374;
176.301252,-144.484443;
176.301252,-133.340374;
176.301252,-150.372674;
190.751679,-150.372674;
190.751679,-138.570553;
190.751679,-144.484443;
-32.549573,-133.340374;
-47.000000,-144.484443;
-47.000000,-133.340374;
-47.000000,-150.372674;
-32.549573,-150.372674;
-32.549573,-144.484443;
-32.549573,-133.340374;
-47.000000,-144.484443;
-47.000000,-133.340374;
-47.000000,-150.372674;
-32.549573,-150.372674;
-32.549573,-144.484443;
-32.549573,-133.340374;
-47.000000,-144.484443;
-47.000000,-133.340374;
-47.000000,-150.372674;
-32.549573,-150.372674;
-32.549573,-144.484443;
-32.549573,-133.340374;
-47.000000,-144.484443;
-47.000000,-133.340374;
-47.000000,-150.372674;
-32.549573,-150.372674;
-32.549573,-144.484443;
-32.549573,-133.340374;
-47.000000,-144.484443;
-47.000000,-133.340374;
-47.000000,-150.372674;
-32.549573,-150.372674;
-32.549573,-144.484443;
190.751679,-133.340374;
176.301252,-144.484443;
176.301252,-133.340374;
176.301252,-150.372674;
190.751679,-150.372674;
190.751679,-144.484443;
190.751679,-133.340374;
176.301252,-144.484443;
176.301252,-133.340374;
176.301252,-150.372674;
190.751679,-150.372674;
190.751679,-144.484443;
190.751679,-133.340374;
176.301252,-144.484443;
176.301252,-133.340374;
176.301252,-150.372674;
190.751679,-150.372674;
190.751679,-144.484443;
190.751679,-133.340374;
176.301252,-144.484443;
176.301252,-133.340374;
176.301252,-150.372674;
190.751679,-150.372674;
190.751679,-144.484443;
190.751679,-133.340374;
176.301252,-144.484443;
176.301252,-133.340374;
176.301252,-150.372674;
190.751679,-150.372674;
190.751679,-144.484443;
190.751679,-133.340374;
176.301252,-144.484443;
176.301252,-133.340374;
176.301252,-150.372674;
190.751679,-150.372674;
190.751679,-144.484443;
190.751679,-133.340374;
176.301252,-144.484443;
176.301252,-133.340374;
176.301252,-150.372674;
190.751679,-150.372674;
190.751679,-144.484443;
-32.549573,-133.340374;
-47.000000,-144.484443;
-47.000000,-133.340374;
-47.000000,-159.987966;
-32.549573,-159.987966;
-32.549573,-144.484443;
190.751679,-133.340374;
176.301252,-144.484443;
176.301252,-133.340374;
176.301252,-159.987966;
190.751679,-159.987966;
190.751679,-144.484443;
-1695.064596,-774.602962;
-1712.499557,-837.900190;
-1962.766740,-700.865507;
-1712.499558,-837.900192;
-1977.550019,-754.535863;
-1962.766741,-700.865508;
-1695.064593,-774.602966;
-1712.499555,-837.900194;
-1962.766738,-700.865510;
-1712.499555,-837.900194;
-1977.550016,-754.535865;
-1962.766737,-700.865510;
-1695.064595,-774.602964;
-1712.499556,-837.900192;
-1962.766739,-700.865509;
-1712.499556,-837.900193;
-1977.550018,-754.535864;
-1962.766739,-700.865509;
-1695.064595,-774.602963;
-1712.499557,-837.900191;
-1962.766739,-700.865508;
-1712.499557,-837.900192;
-1977.550019,-754.535864;
-1962.766740,-700.865508;
-1695.064594,-774.602965;
-1712.499555,-837.900193;
-1962.766738,-700.865509;
-1712.499556,-837.900193;
-1977.550017,-754.535865;
-1962.766738,-700.865510;
-1695.064596,-774.602963;
-1712.499557,-837.900191;
-1962.766740,-700.865507;
-1712.499557,-837.900192;
-1977.550019,-754.535863;
-1962.766740,-700.865508;
-1695.064595,-774.602964;
-1712.499556,-837.900192;
-1962.766739,-700.865508;
-1712.499557,-837.900192;
-1977.550018,-754.535864;
-1962.766739,-700.865509;
-1712.499556,-837.900193;
-1977.550017,-754.535864;
-1962.766739,-700.865509;
-1695.064594,-774.602965;
-1712.499556,-837.900193;
-1962.766739,-700.865509;
-1695.064594,-774.602965;
-1712.499555,-837.900194;
-1962.766738,-700.865510;
-1712.499555,-837.900194;
-1977.550017,-754.535865;
-1962.766738,-700.865510;
-1695.064596,-774.602963;
-1712.499557,-837.900191;
-1962.766740,-700.865507;
-1712.499558,-837.900192;
-1977.550019,-754.535863;
-1962.766740,-700.865508;
-1695.064595,-774.602963;
-1712.499557,-837.900191;
-1962.766740,-700.865507;
-1712.499557,-837.900192;
-1977.550019,-754.535863;
-1962.766740,-700.865508;
-1695.064595,-774.602963;
-1712.499556,-837.900192;
-1962.766739,-700.865508;
-1712.499557,-837.900192;
-1977.550018,-754.535864;
-1962.766740,-700.865509;
-1712.499557,-837.900193;
-1977.550018,-754.535864;
-1962.766739,-700.865509;
-1695.064595,-774.602964;
-1712.499556,-837.900192;
-1962.766739,-700.865508;
-1712.499556,-837.900193;
-1977.550018,-754.535864;
-1962.766739,-700.865509;
-1695.064595,-774.602964;
-1712.499556,-837.900192;
-1962.766739,-700.865509;
-1695.064594,-774.602965;
-1712.499556,-837.900193;
-1962.766738,-700.865509;
-1712.499556,-837.900193;
-1977.550017,-754.535865;
-1962.766738,-700.865509;
-1695.064594,-774.602965;
-1712.499555,-837.900193;
-1962.766738,-700.865510;
-1712.499555,-837.900193;
-1977.550017,-754.535865;
-1962.766738,-700.865510;
-1695.064594,-774.602966;
-1712.499555,-837.900194;
-1962.766738,-700.865510;
-1712.499555,-837.900194;
-1977.550016,-754.535865;
-1962.766738,-700.865510;
-1695.064594,-774.602966;
-1712.499555,-837.900194;
-1962.766738,-700.865510;
-1712.499555,-837.900194;
-1977.550016,-754.535865;
-1962.766738,-700.865510;
-1695.064594,-774.602966;
-1712.499555,-837.900194;
-1962.766738,-700.865510;
-1712.499555,-837.900194;
-1977.550017,-754.535865;
-1962.766738,-700.865510;
-1695.064594,-774.602965;
-1712.499555,-837.900193;
-1962.766738,-700.865510;
-1712.499555,-837.900193;
-1977.550017,-754.535865;
-1962.766738,-700.865510;
-1695.064594,-774.602965;
-1712.499555,-837.900193;
-1962.766738,-700.865510;
-1712.499555,-837.900193;
-1977.550017,-754.535865;
-1962.766738,-700.865510;
-1712.499556,-837.900193;
-1977.550017,-754.535865;
-1962.766738,-700.865509;
-1695.064594,-774.602965;
-1712.499556,-837.900193;
-1962.766738,-700.865509;
-1695.064594,-774.602965;
-1712.499556,-837.900193;
-1962.766738,-700.865509;
-1712.499556,-837.900193;
-1977.550017,-754.535864;
-1962.766739,-700.865509;
-1695.064594,-774.602964;
-1712.499556,-837.900193;
-1962.766739,-700.865509;
-1712.499556,-837.900193;
-1977.550018,-754.535864;
-1962.766739,-700.865509;
-1695.064595,-774.602964;
-1712.499556,-837.900192;
-1962.766739,-700.865509;
-1712.499556,-837.900193;
-1977.550018,-754.535864;
-1962.766739,-700.865509;
-1695.064595,-774.602964;
-1712.499556,-837.900192;
-1962.766739,-700.865508;
-1712.499556,-837.900193;
-1977.550018,-754.535864;
-1962.766739,-700.865509;
-1695.064595,-774.602964;
-1712.499556,-837.900192;
-1962.766739,-700.865508;
-1712.499557,-837.900193;
-1977.550018,-754.535864;
-1962.766739,-700.865509;
-1695.064595,-774.602964;
-1712.499556,-837.900192;
-1962.766739,-700.865508;
-1712.499557,-837.900192;
-1977.550018,-754.535864;
-1962.766740,-700.865509;
-1695.064595,-774.602963;
-1712.499557,-837.900191;
-1962.766739,-700.865508;
-1712.499557,-837.900192;
-1977.550019,-754.535864;
-1962.766740,-700.865509;
-1695.064595,-774.602963;
-1712.499557,-837.900191;
-1962.766739,-700.865508;
-1712.499557,-837.900192;
-1977.550019,-754.535863;
-1962.766740,-700.865508;
-1695.064595,-774.602963;
-1712.499557,-837.900191;
-1962.766740,-700.865507;
-1712.499557,-837.900192;
-1977.550019,-754.535863;
-1962.766740,-700.865508;
-1695.064596,-774.602963;
-1712.499557,-837.900191;
-1962.766740,-700.865507;
-1712.499558,-837.900192;
-1977.550019,-754.535863;
-1962.766740,-700.865508;
-1695.064596,-774.602962;
-1712.499557,-837.900191;
-1962.766740,-700.865507;
-1712.499558,-837.900192;
-1977.550019,-754.535863;
-1962.766741,-700.865508;;
}
MeshMaterialList {
1;
131;
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0;
Material {
1.000000;1.000000;1.000000;1.000000;;
50.000000;
0.000000;0.000000;0.000000;;
0.000000;0.000000;0.000000;;
}
}
}
}
}