#include <math.h>
#include <stdio.h> 


struct Point {
    float X;
    float Y;
    float Z;
};

const double epsilon = 1e-9;

extern bool Test (unsigned short n1, unsigned short n2, unsigned short n3);
/* Accepts inputs n1 and n3 from PointInsidePointList() (two 2D views of one 3D set of data) and n2 from TriangleCrossSegment() (a bridge to skip a third 2D view) */

extern bool PointTouchesTriangle(float pointX, float pointY, float pointZ, float normalX, float normalY, float normalZ, float centerX, float centerY, float centerZ);
/* Checks for the presence of a point possibly behind a triangle, the first three inputs are the point to test with
the triangles center removed, the next three are the triangles normal, the last three are tthe triangles center. */

extern int PointInsidePointList(float pointX, float pointY,float *pointListX, float *pointListY, int pointListCount);
/* Tests for the presence of a 2D point pointX,pointY anywhere within a 2D shape defined with a list of points pointListX,pointListY that has pointListCount number of coordinates,
returning the the unsigned percentage of maximum datatype numerical relation to percentage of total coordinates, or zero if the point does not occur within the shapes defined boundaries. */

extern float TriangleCrossSegment(float Ax1, float Ay1, float Az1, float Ax2, float Ay2, float Az2, float Ax3, float Ay3, float Az3, float Bx1, float By1, float Bz1, float Bx2, float By2, float Bz2, float Bx3, float By3, float Bz3);
extern float TriangleCrossSegmentEx(float Ax1, float Ay1, float Az1, float Ax2, float Ay2, float Az2, float Ax3, float Ay3, float Az3, float Bx1, float By1, float Bz1, float Bx2, float By2, float Bz2, float Bx3, float By3, float Bz3, float *Px0, float *Py0, float *Pz0, float *Px1, float *Py1, float *Pz1);
/* Accepts two trianlges, A, and B, by 3 verticies each, and returns the length of the overlapping segment
line formed from their collision, as well the points of the line segment if the extended version */






extern void CollisionClearFlag (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[]);
/* Resets all flags to Flag of Triangle data, */




extern int CollisionObjectFlag (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngObjectIndex);
/* Resets flags of Traingle data flags to Flag whose object matches lngObjectIndex, returns the number of triangles changed */

extern void CollisionTriangleFlag (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngTriangleIndex, int lngTriangleCount);
/* Resets lngTriangleCount number of Traingle data flags to Flag, starting at lngTriangleIndex */

extern int CollisionResetFlag (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngNewFlag);
/* Resets all flags to lngNewFlag of Triangle data whose flags matches Flag exactly, returns the number of triangles changed */




extern int CollisionObjectCulling (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngObjectIndex, int lngCullingMethod);
/* Culling function that sets Triangle data whose object index is lngObjectIndex to Flag based on a method lngCullingMethod of selecting and/or eleminating non near collision traingles, retruns the number of traingles reduced by */

extern int CollisionTriangleCulling (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngTriangleIndex, int lngTriangleCount, int lngCullingMethod);
/* Culling function that sets lngTraingleCOunt number of Triangle data sarting with lngTriangleIndex to Flag based on a method lngCullingMethod of selecting and/or eleminating non near collision traingles, retruns the number of traingles reduced by */

extern int CollisionFlagCulling (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngNewFlag, int lngCullingMethod);
/* Culling function that sets Triangle data whose flag matches Flag to lngNewFlag based on a method lngCullingMethod of selecting and/or eleminating non near collision traingles, retruns the number of traingles reduced by */




extern bool CollisionChecking (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngTriangleIndex, int *lngCollidedObject, int *lngCollidedTriangle);
/* Tests collision of a the Trianlge data at lngTriangleIndex to all Traingle data whose flags match Flag returning the first instance of collision by setting lngCollidedTriangle of lngCollidedObject,returns true if a collision occurs  */

extern bool Test (unsigned short n1, unsigned short n2, unsigned short n3)
{
	
	return  (bool) ((((n1 & n2 + n3) || (n1 + n2 && n3)) && ((n1 - n2 || !n3) - (!n1 || n2 - n3)))
				 || (((n1 - n2 || n3) && (n1 - n2 || n3)) + ((n1 || n2 + !n3) && (!n1 + n2 && n3))));
}

extern bool PointTouchesTriangle (float pointX, float pointY, float pointZ, float normalX, float normalY, float normalZ, float centerX, float centerY, float centerZ) 
{
	return ((((normalX * pointX) + (normalY * pointY) + (normalZ * pointZ)) - ((normalX * centerX) + (normalY * centerY) + (normalZ * centerZ))) <= 0.0);
}

extern int PointInsidePointList(float pointX, float pointY, float pointListX[], float pointListY[], int pointListCount)
{
	if (pointListCount>2) {
		float ref=((pointX - pointListX[0]) * (pointListY[1] - pointListY[0]) - (pointY - pointListY[0]) * (pointListX[1] - pointListX[0]));
		float ret=ref;
		int result=0;
		for (int i=1;i<=pointListCount;i++) {
			ref = ((pointX - pointListX[i-1]) * (pointListY[i] - pointListY[i-1]) - (pointY - pointListY[i-1]) * (pointListX[i] - pointListX[i-1]));
			if ((ret >= 0) && (ref < 0) && (result==0)) result = i;
			ret=ref;
		}
		if ((result==0)||(result>pointListCount)) return 1;
	}
	return 0;
}

extern void CollisionClearFlag (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[]) {
/* Resets all flags to Flag of Triangle data, */
	for (int i=0; i<lngTriangleTotal; i++) {
		sngFaceVis[3][i] = 0;
	}
}

extern int CollisionObjectFlag (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngObjectIndex) {
/* Resets flags of Traingle data flags to Flag whose object matches lngObjectIndex */
	int cnt;
	for (int i=0; i<lngTriangleTotal; i++) {
		if (sngFaceVis[4][i]==lngObjectIndex) {
			sngFaceVis[3][i] = Flag;
			cnt++;			
		}
	}
	return cnt;
}

extern void CollisionTriangleFlag (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngTriangleIndex, int lngTriangleCount) {
/* Resets lngTriangleCount number of Traingle data flags to Flag, starting at lngTriangleIndex */
	for (int i=lngTriangleIndex; i<(lngTriangleIndex+lngTriangleCount); i++) {
		sngFaceVis[3][i] = Flag;
	}
}
extern int CollisionResetFlag (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngNewFlag) {
/* Resets all flags to lngNewFlag of Triangle data whose flags matches Flag exactly, */
	int cnt;
	for (int i=0; i<lngTriangleTotal; i++) {
		if (sngFaceVis[3][i] == Flag) {
			sngFaceVis[3][i]=lngNewFlag;
			cnt++;
		}
	}
	return cnt;
}


extern int CollisionObjectCulling (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngObjectIndex, int lngCullingMethod) {
/* Culling function that sets Triangle data whose object index is lngObjectIndex to Flag based on a method lngCullingMethod of selecting and/or eleminating non near collision traingles */

	return 0;
}
extern int CollisionTriangleCulling (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngTriangleIndex, int lngTriangleCount, int lngCullingMethod) {
/* Culling function that sets lngTraingleCOunt number of Triangle data sarting with lngTriangleIndex to Flag based on a method lngCullingMethod of selecting and/or eleminating non near collision traingles */

	return 0;
}
extern int CollisionFlagCulling (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngNewFlag, int lngCullingMethod) {
/* Culling function that sets Triangle data whose flag matches Flag to lngNewFlag based on a method lngCullingMethod of selecting and/or eleminating non near collision traingles */
	return 0;
}

extern int CollisionCulling /* was Forystek */ (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[])
{
	//todo, any awesomer culling that can be done which only applies if falls in a options flag manor defined by the user so multiple calls can paint the canvas of 
	//from single player with their bullets to online load balancing in systems processing potentially quicker when objects and their counter parts hit the scenery.
	//this function crashed at vistype>3 and it also didn't only flag the applied, it reset every flag on every call so collective map flagging was not poossible.
	return 0;
}
extern bool CollisionChecking (int Flag, int lngTriangleTotal, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngFaceNum, int *lngCollidedBrush, int *lngCollidedFace)
{

					 


	return 0;
}





// ===== Helpers =====

Point MakePoint(float x, float y, float z) {
    Point p; p.X = x; p.Y = y; p.Z = z; return p;
}

Point VectorAddition(Point a, Point b) {
    Point r; r.X = b.X + a.X; r.Y = b.Y +  a.Y; r.Z = b.Z + a.Z; return r;
}

Point VectorDeduction(Point a, Point b) {
    Point r; r.X = a.X - b.X; r.Y = a.Y - b.Y; r.Z = a.Z - b.Z; return r;
}

float VectorDotProduct(Point a, Point b) {
    return a.X*b.X + a.Y*b.Y + a.Z*b.Z;
}

Point VectorCrossProduct(Point a, Point b) {
    Point r;
    r.X = a.Y*b.Z - a.Z*b.Y;
    r.Y = a.Z*b.X - a.X*b.Z;
    r.Z = a.X*b.Y - a.Y*b.X;
    return r;
}

float Distance(Point p1, Point p2) {
    float dx = p1.X - p2.X;
    float dy = p1.Y - p2.Y;
    float dz = p1.Z - p2.Z;
    float sumSq = dx*dx + dy*dy + dz*dz;
    return (sumSq != 0.0) ? sqrtf(sumSq) : 0;
}


Point TriangleNormal(Point p1, Point p2, Point p3) {
    Point v1 = VectorDeduction(p2, p1);
    Point v2 = VectorDeduction(p3, p1);
    return VectorCrossProduct(v1, v2);
}

float Length(Point p) {
    return sqrtf(p.X*p.X + p.Y*p.Y + p.Z*p.Z);
}

Point VectorNormalize(Point p) {
    float len = Length(p);
    if (len == 0.0) return MakePoint(0,0,0);
    return MakePoint(p.X/len, p.Y/len, p.Z/len);
}

float Least(float a, float b) { return (a < b) ? a : b; }
float Least(float a, float b, float c) { return (a < b) ? (a < c) ? a : c : (b < c) ? b : c; }
float Large(float a, float b) { return (a > b) ? a : b; }
float Large(float a, float b, float c) { return (a > b) ? (a > c) ? a : c : (b > c) ? b : c; }

// ===== Geometry checks =====
int AreParallel(Point t1p1, Point t1p2, Point t1p3,
                Point t2p1, Point t2p2, Point t2p3) {
    //const double EPS = 1e-9;
    Point n1 = TriangleNormal(t1p1, t1p2, t1p3);
    Point n2 = TriangleNormal(t2p1, t2p2, t2p3);
    Point cross = VectorCrossProduct(n1, n2);
    return (fabsf(cross.X) < epsilon &&
            fabsf(cross.Y) < epsilon &&
            fabsf(cross.Z) < epsilon);
}

int AreCoplanar(Point t1p1, Point t1p2, Point t1p3,
                Point t2p1, Point t2p2, Point t2p3) {
    
    if (!AreParallel(t1p1, t1p2, t1p3, t2p1, t2p2, t2p3)) return 0;
    Point n1 = TriangleNormal(t1p1, t1p2, t1p3);
    float d = -(n1.X*t1p1.X + n1.Y*t1p1.Y + n1.Z*t1p1.Z);
    float planeEq = n1.X*t2p1.X + n1.Y*t2p1.Y + n1.Z*t2p1.Z + d;
    return (fabsf(planeEq) < epsilon);
}

int PointInTriangle(Point p, Point V0, Point v1, Point v2) {
    //const double EPS = 1e-9;
    Point u = VectorDeduction(v1, V0);
    Point v = VectorDeduction(v2, V0);
    Point w = VectorDeduction(p, V0);

    float uu = VectorDotProduct(u,u);
    float vv = VectorDotProduct(v,v);
    float uv = VectorDotProduct(u,v);
    float wu = VectorDotProduct(w,u);
    float wv = VectorDotProduct(w,v);

    float d = uv*uv - uu*vv;
    if (fabs(d) < epsilon) return 0;

    float s = (uv*wv - vv*wu) / d;
    float t = (uv*wu - uu*wv) / d;

    return (s >= -epsilon && t >= -epsilon && (s+t) <= 1.0+epsilon);
}

int EdgePlaneIntersect(Point p, Point Q, Point planePoint, Point PlaneNormal, Point &X) {
    //const double EPS = 1e-9;
    Point dir = VectorDeduction(Q, p);
    float denom = VectorDotProduct(PlaneNormal, dir);
    if (fabs(denom) < epsilon) return 0;

    float t = VectorDotProduct(PlaneNormal, VectorDeduction(planePoint, p)) / denom;
    if (t < -epsilon || t > 1.0+epsilon) return 0;

    X = VectorAddition(p, MakePoint(dir.X*t, dir.Y*t, dir.Z*t));
    return 1;
}


//inputs: verticies for two triangles in collision (defined by two or more axis of PointInPoly 2D collision tests)
//output: two points to form a line segment where the trianlges collide.
//retruns: length of overlap segment
extern float TriangleCrossSegment(float Ax1, float Ay1, float Az1, float Ax2, float Ay2, float Az2, float Ax3, float Ay3, float Az3, 
						   float Bx1, float By1, float Bz1, float Bx2, float By2, float Bz2, float Bx3, float By3, float Bz3)
{
	float *Px0=0;float *Py0=0;float *Pz0=0;float *Px1=0;float *Py1=0;float *Pz1=0;
	return TriangleCrossSegmentEx(Ax1, Ay1, Az1, Ax2, Ay2, Az2, Ax3, Ay3, Az3, 
							Bx1, By1, Bz1, Bx2, By2, Bz2, Bx3, By3, Bz3,
							Px0, Py0, Pz0, Px1, Py1, Pz1);

}
extern float TriangleCrossSegmentEx(float Ax1, float Ay1, float Az1, float Ax2, float Ay2, float Az2, float Ax3, float Ay3, float Az3, 
						   float Bx1, float By1, float Bz1, float Bx2, float By2, float Bz2, float Bx3, float By3, float Bz3, 
						   float *Px0, float *Py0, float *Pz0, float *Px1, float *Py1, float *Pz1)
{
	Point t1p1=MakePoint(Ax1, Ay1, Az1);
	Point t1p2=MakePoint(Ax2, Ay2, Az2);
	Point t1p3=MakePoint(Ax3, Ay3, Az3);
	Point t2p1=MakePoint(Bx1, By1, Bz1);
	Point t2p2=MakePoint(Bx2, By2, Bz2);
	Point t2p3=MakePoint(Bx3, By3, Bz3);

    int ap = AreParallel(t1p1, t1p2, t1p3, t2p1, t2p2, t2p3);
    int ac = AreCoplanar(t1p1, t1p2, t1p3, t2p1, t2p2, t2p3);

    float l1, l2;

    if (ap && !ac) {
        // Parallel but not coplanar
        return 0.0;
    } else if (ac) {
        // Coplanar case
        l1 = Distance(t1p1,t1p2) + Distance(t1p2,t1p3) + Distance(t1p3,t1p1);
        l2 = Distance(t2p1,t2p2) + Distance(t2p2,t2p3) + Distance(t2p3,t2p1);
        return Least(l1,l2);//the smallest is the whole overlap all the wasy around
    } else {
        // Intersecting, non-coplanar triangles
        Point nA = VectorCrossProduct(VectorDeduction(t1p2,t1p1), VectorDeduction(t1p3,t1p1));
        Point nB = VectorCrossProduct(VectorDeduction(t2p2,t2p1), VectorDeduction(t2p3,t2p1));

        Point pts[6];
        int C = 0;
        Point X;

        // Intersect edges of A with plane of B
        if (EdgePlaneIntersect(t1p1,t1p2,t2p1,nB,X) && PointInTriangle(X,t2p1,t2p2,t2p3)) pts[C++] = X;
        if (EdgePlaneIntersect(t1p2,t1p3,t2p1,nB,X) && PointInTriangle(X,t2p1,t2p2,t2p3)) pts[C++] = X;
        if (EdgePlaneIntersect(t1p3,t1p1,t2p1,nB,X) && PointInTriangle(X,t2p1,t2p2,t2p3)) pts[C++] = X;

        // Intersect edges of B with plane of A
        if (EdgePlaneIntersect(t2p1,t2p2,t1p1,nA,X) && PointInTriangle(X,t1p1,t1p2,t1p3)) pts[C++] = X;
        if (EdgePlaneIntersect(t2p2,t2p3,t1p1,nA,X) && PointInTriangle(X,t1p1,t1p2,t1p3)) pts[C++] = X;
        if (EdgePlaneIntersect(t2p3,t2p1,t1p1,nA,X) && PointInTriangle(X,t1p1,t1p2,t1p3)) pts[C++] = X;

        if (C < 2) {
            // Shouldn’t happen if collision preconditions are met
            return 0.0;
        } else {

			// Choose two extreme points along intersection line direction
			Point dir = VectorNormalize(VectorCrossProduct(nA,nB));

			float minProj = VectorDotProduct(dir, pts[0]);
			float maxProj = minProj;
			int minIdx = 0, maxIdx = 0;

			for (int i=1; i<C; i++) {
				float p = VectorDotProduct(dir, pts[i]);
				if (p < minProj) { minProj = p; minIdx = i; }
				if (p > maxProj) { maxProj = p; maxIdx = i; }
			}

			*Px0 = pts[minIdx].X;
			*Py0 = pts[minIdx].Y;
			*Pz0 = pts[minIdx].Z;

			*Px1 = pts[maxIdx].X;
			*Py1 = pts[maxIdx].Y;
			*Pz1 = pts[maxIdx].Z;

			return Distance(MakePoint(*Px0,*Py0,*Pz0),MakePoint(*Px1,*Py1,*Pz1));
		}
	}

}

 