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
96;
0.000000;0.000000;-0.304800;,
0.000000;-0.215526;-0.215526;,
0.215526;0.000000;-0.215526;,
-0.215526;0.000000;0.215526;,
-0.215526;-0.215526;0.000000;,
-0.304800;0.000000;0.000000;,
0.215526;0.000000;0.215526;,
0.215526;-0.215526;0.000000;,
0.000000;-0.215526;0.215526;,
-0.304800;0.000000;0.000000;,
-0.215526;0.000000;-0.215526;,
-0.215526;0.215526;0.000000;,
0.000000;0.215526;-0.215526;,
-0.215526;0.215526;0.000000;,
-0.215526;0.000000;-0.215526;,
-0.215526;0.000000;0.215526;,
-0.304800;0.000000;0.000000;,
-0.215526;0.215526;0.000000;,
0.000000;-0.215526;-0.215526;,
-0.215526;-0.215526;0.000000;,
0.000000;-0.304800;0.000000;,
0.000000;0.304800;0.000000;,
-0.215526;0.215526;0.000000;,
0.000000;0.215526;-0.215526;,
0.000000;-0.304800;0.000000;,
-0.215526;-0.215526;0.000000;,
0.000000;-0.215526;0.215526;,
0.215526;0.215526;0.000000;,
0.000000;0.304800;0.000000;,
0.000000;0.215526;-0.215526;,
0.215526;-0.215526;0.000000;,
0.000000;-0.304800;0.000000;,
0.000000;-0.215526;0.215526;,
0.215526;0.215526;0.000000;,
0.000000;0.215526;-0.215526;,
0.215526;0.000000;-0.215526;,
0.000000;0.215526;0.215526;,
0.215526;0.000000;0.215526;,
0.000000;0.000000;0.304800;,
0.215526;0.215526;0.000000;,
0.000000;0.215526;0.215526;,
0.000000;0.304800;0.000000;,
-0.215526;-0.215526;0.000000;,
-0.215526;0.000000;-0.215526;,
-0.304800;0.000000;0.000000;,
0.000000;0.000000;0.304800;,
0.000000;-0.215526;0.215526;,
-0.215526;0.000000;0.215526;,
0.215526;-0.215526;0.000000;,
0.215526;0.000000;-0.215526;,
0.000000;-0.215526;-0.215526;,
0.215526;0.000000;0.215526;,
0.304800;0.000000;0.000000;,
0.215526;-0.215526;0.000000;,
0.000000;0.215526;0.215526;,
-0.215526;0.215526;0.000000;,
0.000000;0.304800;0.000000;,
0.000000;0.215526;-0.215526;,
-0.215526;0.000000;-0.215526;,
0.000000;0.000000;-0.304800;,
0.000000;0.215526;-0.215526;,
0.000000;0.000000;-0.304800;,
0.215526;0.000000;-0.215526;,
0.215526;0.000000;0.215526;,
0.215526;0.215526;0.000000;,
0.304800;0.000000;0.000000;,
0.215526;-0.215526;0.000000;,
0.000000;-0.215526;-0.215526;,
0.000000;-0.304800;0.000000;,
0.000000;0.215526;0.215526;,
-0.215526;0.000000;0.215526;,
-0.215526;0.215526;0.000000;,
-0.215526;0.000000;-0.215526;,
0.000000;-0.215526;-0.215526;,
0.000000;0.000000;-0.304800;,
0.304800;0.000000;0.000000;,
0.215526;0.000000;-0.215526;,
0.215526;-0.215526;0.000000;,
-0.215526;0.000000;-0.215526;,
-0.215526;-0.215526;0.000000;,
0.000000;-0.215526;-0.215526;,
0.000000;-0.215526;0.215526;,
-0.215526;-0.215526;0.000000;,
-0.215526;0.000000;0.215526;,
0.215526;0.215526;0.000000;,
0.215526;0.000000;-0.215526;,
0.304800;0.000000;0.000000;,
0.215526;0.000000;0.215526;,
0.000000;-0.215526;0.215526;,
0.000000;0.000000;0.304800;,
0.000000;0.215526;0.215526;,
0.000000;0.000000;0.304800;,
-0.215526;0.000000;0.215526;,
0.215526;0.000000;0.215526;,
0.000000;0.215526;0.215526;,
0.215526;0.215526;0.000000;;
32;
3;2,1,0,
3;5,4,3,
3;8,7,6,
3;11,10,9,
3;14,13,12,
3;17,16,15,
3;20,19,18,
3;23,22,21,
3;26,25,24,
3;29,28,27,
3;32,31,30,
3;35,34,33,
3;38,37,36,
3;41,40,39,
3;44,43,42,
3;47,46,45,
3;50,49,48,
3;53,52,51,
3;56,55,54,
3;59,58,57,
3;62,61,60,
3;65,64,63,
3;68,67,66,
3;71,70,69,
3;74,73,72,
3;77,76,75,
3;80,79,78,
3;83,82,81,
3;86,85,84,
3;89,88,87,
3;92,91,90,
3;95,94,93;;
MeshNormals {
96;
0.357407;-0.357407;-0.862856;
0.357407;-0.357407;-0.862856;
0.357407;-0.357407;-0.862856;
-0.862856;-0.357407;0.357407;
-0.862856;-0.357407;0.357407;
-0.862856;-0.357407;0.357407;
0.577350;-0.577350;0.577350;
0.577350;-0.577350;0.577350;
0.577350;-0.577350;0.577350;
-0.862856;0.357407;-0.357407;
-0.862856;0.357407;-0.357407;
-0.862856;0.357407;-0.357407;
-0.577350;0.577350;-0.577350;
-0.577350;0.577350;-0.577350;
-0.577350;0.577350;-0.577350;
-0.862856;0.357407;0.357407;
-0.862856;0.357407;0.357407;
-0.862856;0.357407;0.357407;
-0.357407;-0.862856;-0.357407;
-0.357407;-0.862856;-0.357407;
-0.357407;-0.862856;-0.357407;
-0.357407;0.862856;-0.357407;
-0.357407;0.862856;-0.357407;
-0.357407;0.862856;-0.357407;
-0.357407;-0.862856;0.357407;
-0.357407;-0.862856;0.357407;
-0.357407;-0.862856;0.357407;
0.357407;0.862856;-0.357407;
0.357407;0.862856;-0.357407;
0.357407;0.862856;-0.357407;
0.357407;-0.862856;0.357407;
0.357407;-0.862856;0.357407;
0.357407;-0.862856;0.357407;
0.577350;0.577350;-0.577350;
0.577350;0.577350;-0.577350;
0.577350;0.577350;-0.577350;
0.357407;0.357407;0.862856;
0.357407;0.357407;0.862856;
0.357407;0.357407;0.862856;
0.357407;0.862856;0.357407;
0.357407;0.862856;0.357407;
0.357407;0.862856;0.357407;
-0.862856;-0.357407;-0.357407;
-0.862856;-0.357407;-0.357407;
-0.862856;-0.357407;-0.357407;
-0.357407;-0.357407;0.862856;
-0.357407;-0.357407;0.862856;
-0.357407;-0.357407;0.862856;
0.577350;-0.577350;-0.577350;
0.577350;-0.577350;-0.577350;
0.577350;-0.577350;-0.577350;
0.862856;-0.357407;0.357407;
0.862856;-0.357407;0.357407;
0.862856;-0.357407;0.357407;
-0.357407;0.862856;0.357407;
-0.357407;0.862856;0.357407;
-0.357407;0.862856;0.357407;
-0.357407;0.357407;-0.862856;
-0.357407;0.357407;-0.862856;
-0.357407;0.357407;-0.862856;
0.357407;0.357407;-0.862856;
0.357407;0.357407;-0.862856;
0.357407;0.357407;-0.862856;
0.862856;0.357407;0.357407;
0.862856;0.357407;0.357407;
0.862856;0.357407;0.357407;
0.357407;-0.862856;-0.357407;
0.357407;-0.862856;-0.357407;
0.357407;-0.862856;-0.357407;
-0.577350;0.577350;0.577350;
-0.577350;0.577350;0.577350;
-0.577350;0.577350;0.577350;
-0.357407;-0.357407;-0.862856;
-0.357407;-0.357407;-0.862856;
-0.357407;-0.357407;-0.862856;
0.862856;-0.357407;-0.357407;
0.862856;-0.357407;-0.357407;
0.862856;-0.357407;-0.357407;
-0.577350;-0.577350;-0.577350;
-0.577350;-0.577350;-0.577350;
-0.577350;-0.577350;-0.577350;
-0.577350;-0.577350;0.577350;
-0.577350;-0.577350;0.577350;
-0.577350;-0.577350;0.577350;
0.862856;0.357407;-0.357407;
0.862856;0.357407;-0.357407;
0.862856;0.357407;-0.357407;
0.357407;-0.357407;0.862856;
0.357407;-0.357407;0.862856;
0.357407;-0.357407;0.862856;
-0.357407;0.357407;0.862856;
-0.357407;0.357407;0.862856;
-0.357407;0.357407;0.862856;
0.577350;0.577350;0.577350;
0.577350;0.577350;0.577350;
0.577350;0.577350;0.577350;;
32;
3;2,1,0;
3;5,4,3;
3;8,7,6;
3;11,10,9;
3;14,13,12;
3;17,16,15;
3;20,19,18;
3;23,22,21;
3;26,25,24;
3;29,28,27;
3;32,31,30;
3;35,34,33;
3;38,37,36;
3;41,40,39;
3;44,43,42;
3;47,46,45;
3;50,49,48;
3;53,52,51;
3;56,55,54;
3;59,58,57;
3;62,61,60;
3;65,64,63;
3;68,67,66;
3;71,70,69;
3;74,73,72;
3;77,76,75;
3;80,79,78;
3;83,82,81;
3;86,85,84;
3;89,88,87;
3;92,91,90;
3;95,94,93;;
}
MeshTextureCoords {
96;
0.695660,0.189425;
0.696781,0.196996;
0.703313,0.189425;
0.827370,-0.215528;
0.833903,-0.207957;
0.835024,-0.215528;
1.613192,0.035381;
1.608192,0.044041;
1.618192,0.044041;
0.456936,0.110708;
0.464590,0.110708;
0.458057,0.103136;
0.391808,0.038267;
0.381808,0.038267;
0.386808,0.046928;
0.827370,0.215528;
0.835024,0.215528;
0.833903,0.207957;
0.391808,-0.062246;
0.381808,-0.062246;
0.386808,-0.056451;
0.386808,0.056451;
0.381808,0.062246;
0.391808,0.062246;
1.071281,-0.524042;
1.076281,-0.529836;
1.066281,-0.529836;
0.933719,-0.528356;
0.928719,-0.534151;
0.923719,-0.528356;
1.608192,0.060765;
1.613192,0.066560;
1.618192,0.060765;
0.933719,-0.356913;
0.923719,-0.356913;
0.928719,-0.348253;
1.596499,0.056062;
1.589966,0.063633;
1.597620,0.063633;
1.608192,-0.060765;
1.618192,-0.060765;
1.613192,-0.066560;
0.458057,-0.103136;
0.464590,-0.110708;
0.456936,-0.110708;
1.296687,-0.196029;
1.297807,-0.188458;
1.304340,-0.196029;
0.933719,0.356913;
0.928719,0.348253;
0.923719,0.356913;
1.543064,0.104104;
1.535410,0.104104;
1.536531,0.111675;
1.066281,0.529836;
1.076281,0.529836;
1.071281,0.524042;
0.408913,-0.064601;
0.402380,-0.057029;
0.410034,-0.057029;
0.696781,-0.196996;
0.695660,-0.189425;
0.703313,-0.189425;
1.543064,-0.104104;
1.536531,-0.111675;
1.535410,-0.104104;
0.933719,0.528356;
0.923719,0.528356;
0.928719,0.534151;
1.066281,0.351140;
1.071281,0.359800;
1.076281,0.351140;
0.402380,0.057029;
0.408913,0.064601;
0.410034,0.057029;
1.172630,0.208924;
1.164976,0.208924;
1.171509,0.216495;
0.386808,-0.046928;
0.381808,-0.038267;
0.391808,-0.038267;
1.066281,-0.351140;
1.076281,-0.351140;
1.071281,-0.359800;
1.171509,-0.216495;
1.164976,-0.208924;
1.172630,-0.208924;
1.589966,-0.063633;
1.596499,-0.056062;
1.597620,-0.063633;
1.297807,0.188458;
1.296687,0.196029;
1.304340,0.196029;
1.613192,-0.035381;
1.618192,-0.044041;
1.608192,-0.044041;;
}
MeshMaterialList {
2;
32;
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1;
Material {
1.000000;1.000000;1.000000;1.000000;;
50.000000;
0.000000;0.000000;0.000000;;
0.000000;0.000000;0.000000;;
}
Material {
0.501961;0.141176;0.356863;1.000000;;
50.000000;
0.000000;0.000000;0.000000;;
0.000000;0.000000;0.000000;;
TextureFilename { "bubble2.bmp"; }
}
}
}
}
}