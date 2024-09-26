#include <math.h>
#include <stdio.h> 


//Due to some code loss, this DLL will not function as the MaxLandLib.dll used in the gaming collision from Neotext
//I am merely trying to recouperate a DLL as functional as with the same performance but it is taking a long time.
//should these be correct so far, then the three hardest functions, or longest are yet to come, and decompile may
//satisfy, I'm not so sure yet.  I don't ever program C++ and so the default project save folder was not backed up.
//
//Test() so far is producing same results as the already compiled binary MaxLandLib.dll's Test()
//PointBehindPoly() was from IDA Pro Hex-Rays, an online version at some other's website
//PointInPoly() I don't remember the state of texting results compared to MaxLandLib.dll's but
//the general idea of the function is there if only it were done correct, and returns fee to Test()
//tri_tri_intersect() also returns feed to Test(), together with PointInPoly() and I plan on
//renaming these functions when I have them all better back to themselves with code, I even plan
//on putting in DLL file information, but I would like the new code to be fully functioning in
//the current games that use it for collision before I do change the interface.
//

extern bool Test (unsigned short n1, unsigned short n2, unsigned short n3);
/* Accepts inputs n1 and n2 as retruned from PointInPoly(X,Y) then again for (Z,Y) and n2 as returned from tri_tri_intersect() to return the determination of whether or not the collision is correct and satisfy bitwise and math equalaterally collision precise to real coordination from the preliminary possible collision information the other functions return. */

extern bool PointBehindPoly (float a1, float a2, float a3, float a4, float a5, float a6, float a7, float a8, float a9);
/* Checks for the presence of a point behind a triangle, the first three inputs are the length of the triangles sides, the next three are the triangles normal, the last three are the point to test with the triangles center removed. */

extern int PointInPoly ( float pointX, float pointY, float polyDataX[], float polyDataY[], int polyDataCount);
//extern int PointInPoly (float pointX,  float pointY, float polyDataX[], float polyDataY[], int polyDataCount);
//extern int PointInPoly (float testx, float testy, float *vertx, float *verty, int nvert);

//extern short PointInPoly(float pX, float pY,float *polyX, float *polyY, short polyN);
/* Tests for the presence of a 2D point pX,pY anywhere within a 2D shape defined with a list of points polyX,polyY that has polyN number of coordinates, returning the the unsigned percentage of maximum datatype numerical relation to percentage of total coordinates, or zero if the point does not occur within the shapes defined boundaries. */
extern short tri_tri_intersect (unsigned short v0_0, unsigned short v0_1, unsigned short v0_2, unsigned short v1_0, unsigned short v1_1, unsigned short v1_2, unsigned short v2_0, unsigned short v2_1, unsigned short v2_2, unsigned short u0_0, unsigned short u0_1, unsigned short u0_2, unsigned short u1_0, unsigned short u1_1, unsigned short u1_2, unsigned short u2_0, unsigned short u2_1, unsigned short u2_2);
/* Accepts two triangle inputs in hyperbolic paraboloid collision form and returns with in the unsiged whole the percentage of each others distance to plane as one value.  **NOTE Assumes the parameter input as triangles are TRUE for collision with one another. */

extern int Forystek (int visType, int lngFaceCount, unsigned short *sngCamera[], unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], unsigned short *sngScreenX[], unsigned short *sngScreenY[], unsigned short *sngScreenZ[], unsigned short *sngZBuffer[]);
/* Culling function with three expirimental ways to cull, defined by visType, 0 to 2, returns the difference of input triangles. lngFaceCount, sngCamera[3 x 3], sngFaceVis[6 x lngFaceCount], sngVertexX[3 x lngFaceCount]..Y..Z, sngScreenX[3 x lngFaceCount]..Y..Z, sngZBuffer[4 x lngFaceCount].  The camera is defined by position [0,0]=X, [0,1]=Y, [0,2]=Z, direction [1,0]=X, [1,1]=Y, [1,2]=Z, and upvector [2,0]=X, [2,1]=Y, [2,2]=Z.  sngFaceVis should be initialized to zero, and sngVertex arrays are 3D coordinate equivelent to sngScreen with a screenZ buffer, and Zbuffer for the verticies. */

extern bool Collision (int visType, int lngFaceCount, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngFaceNum, int *lngCollidedBrush, int *lngCollidedFace);
/* Tests collision of a lngFaceNum against a number of visible faces, lngFaceCount, whose sngFaceVis has been defined with visType as culled with the Forystek function, and returns whether or not a collision occurs also populating the lngCollidedBrush and lngCollidedFace indicating the exact object number (brush) and face number (triangle) that has the collision impact. */





extern bool Test (unsigned short n1, unsigned short n2, unsigned short n3)
{
	
	return  (bool) ((((n1 && n2 + n3) || (n1 + n2 && n3)) && ((n1 - n2 || !n3) - (!n1 || n2 - n3)))
				 || (((n1 - n2 || n3) && (n1 - n2 || n3)) + ((n1 || n2 + !n3) && (!n1 + n2 && n3))));
}

extern bool PointBehindPoly (float pointX, float pointY, float pointZ, float length1, float length2, float length3, float normalX, float normalY, float normalZ) 
{
	return pointZ * length3 + length2 * pointY + length1 * pointX - (length3 * normalZ + length1 * normalX + length2 * normalY) <= 0.0;
}

extern int PointInPoly ( float pointX, float pointY, float polyDataX[], float polyDataY[], int polyDataCount)
{
	if (polyDataCount>2) {
		float ref=((pointX - polyDataX[0]) * (polyDataY[1] - polyDataY[0]) - (pointY - polyDataY[0]) * (polyDataX[1] - polyDataX[0]));
		float ret=ref;
		int result=0;
		for (int i=1;i<=polyDataCount;i++) {
			ref = ((pointX - polyDataX[i]) * (polyDataY[i] - polyDataY[i-1]) - (pointY - polyDataY[i]) * (polyDataX[i] - polyDataX[i-1]));
			if ((ret > 0) && (ref < 0) && (result==0)) result = i;
			ret=ref;
		}
		if (result!=0) {
			return ((ret>0) && (ref>0));
		} else {
			return ((ret>0) ^ (ref<0));

		}
		//if ((result==0)||(result>polyDataCount)) {
		//	return 1;//todo: this is suppose to return a decimal percent
		//			//of the total polygon points where in is found inside
		//} 
	}
	return 0;
}

/*
extern int PointInPoly ( float pointX, float pointY, float polyDataX[], float polyDataY[], int polyDataCount)
{
	if (polyDataCount>2) {
		float ref=((pointX - polyDataX[0]) * (polyDataY[1] - polyDataY[0]) - (pointY - polyDataY[0]) * (polyDataX[1] - polyDataX[0]));
		float ret=ref;
		int result=0;
		for (int i=1;i<=polyDataCount;i++) {
			ref = ((pointX - polyDataX[i-1]) * (polyDataY[i] - polyDataY[i-1]) - (pointY - polyDataY[i-1]) * (polyDataX[i] - polyDataX[i-1]));
			if ((ret >= 0) && (ref < 0) && (result==0)) result = i;
			ret=ref;
		}
		if ((result==0)||(result>polyDataCount)) return 1;//todo: this is suppose to return a decimal percent
														//of the total polygon points where in is found inside
	}
	return 0;
}
*/
extern short tri_tri_intersect (unsigned short v0_0, unsigned short v0_1, unsigned short v0_2, unsigned short v1_0, unsigned short v1_1, unsigned short v1_2, unsigned short v2_0, unsigned short v2_1, unsigned short v2_2, unsigned short u0_0, unsigned short u0_1, unsigned short u0_2, unsigned short u1_0, unsigned short u1_1, unsigned short u1_2, unsigned short u2_0, unsigned short u2_1, unsigned short u2_2)
{
	return 0;
}

extern int Forystek (int visType, int lngFaceCount, unsigned short *sngCamera[], unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], unsigned short *sngScreenX[], unsigned short *sngScreenY[], unsigned short *sngScreenZ[], unsigned short *sngZBuffer[])
{
	return 0;
}
extern bool Collision (int visType, int lngFaceCount, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngFaceNum, int *lngCollidedBrush, int *lngCollidedFace)
{
	return 0;
}
