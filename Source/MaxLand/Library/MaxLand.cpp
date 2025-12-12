#include <math.h>
#include <stdio.h> 

/*
//the higher the fulling method the more inclusive the triangles are vs apply all culling in potential defaults simply
#define UseAllCulling 0  //apply all the blow methods, culling for collision is the art of painting and weaning flags
#define CullByFlagSet 1  //if culling or flags are used, this is automatic, or else csonder using flag 0 every call
#define CullBySquares 2  //a laments term version of ByCameras, this defined rectangles for three axis that encumbant the triangle and eliminates all trainalges not found with-in the rectangles entirety, has near issues
#define CullByRanging 4  //this finds front faced traingles in it's range and the range is defined by a maximum permiter a traianlge could relate in all three edge lengths, and a spherical "with-in" test to it's center
#define CullByClosest 8  //this is a more refined version of Ranging, simply put, exact cull to one, the closest triangle, after ranging is applied, very effective when scaled down the test by Ranging
#define CullByCameras 16  //a very strong control to approch of culling effective in multiple call projections, applying a Up/Eye/Dir anything with in view, like a rectangle only not so horizontal and vertical locked, becomes included and not culled
#define CullByBehinds 32  //by defualt culling does not include backfacings, unless you include it with this flag, making any triangle having both sides in the matter of culling definition to collision
*/

enum CullingMethod {
	 UseAllCulling =0,
	 CullByFlagSet =1,
	 CullBySquares =2,
	 CullByRanging =4,
	 CullByClosest =8,
	 CullByCameras =16,
	 CullByBehinds =32
};

struct Point {
    float X;
    float Y;
    float Z;
};

const double epsilon = 1e-9;

extern bool Test (unsigned short n1, unsigned short n2, unsigned short n3);
/* Accepts inputs n1 and n3 from PointInsidePointList() (two 2D views of one 3D set of data) and n2 from TriangleCrossSegment() (a bridge to skip a third 2D view) */

extern bool PointTouchesTriangle(float PointX, float PointY, float PointZ, float NormalX, float NormalY, float NormalZ, float CenterX, float CenterY, float CenterZ);
/* Checks for the presence of a point possibly behind a triangle, the first three inputs are the point to test with
the triangles center removed, the next three are the triangles normal, the last three are tthe triangles center. */

extern int PointInsidePointList(float PointX, float PointY,float *PointListX, float *PointListY, int PointListCount);
/* Tests for the presence of a 2D point pointX,pointY anywhere within a 2D shape defined with a list of points pointListX,pointListY that has pointListCount number of coordinates,
returning the the unsigned percentage of maximum datatype numerical relation to percentage of total coordinates, or zero if the point does not occur within the shapes defined boundaries. */

extern float TriangleCrossSegment(float Ax1, float Ay1, float Az1, float Ax2, float Ay2, float Az2, float Ax3, float Ay3, float Az3, float Bx1, float By1, float Bz1, float Bx2, float By2, float Bz2, float Bx3, float By3, float Bz3);
extern float TriangleCrossSegmentEx(float Ax1, float Ay1, float Az1, float Ax2, float Ay2, float Az2, float Ax3, float Ay3, float Az3, float Bx1, float By1, float Bz1, float Bx2, float By2, float Bz2, float Bx3, float By3, float Bz3, float *Px0, float *Py0, float *Pz0, float *Px1, float *Py1, float *Pz1);
/* Accepts two trianlges, A, and B, by 3 verticies each, and returns the length of the overlapping segment
line formed from their collision, as well the points of the line segment if the extended version */


extern void CollisionClearFlag (int Flag, int TriangleTotal, unsigned short *FaceVis[], unsigned short *VertexX[], unsigned short *VertexY[], unsigned short *VertexZ[]);
/* Resets all flags to Flag of Triangle data, */

extern int CollisionObjectFlag (int Flag, int TriangleTotal, unsigned short *FaceVis[], unsigned short *VertexX[], unsigned short *VertexY[], unsigned short *VertexZ[], int ObjectIndex);
/* Resets flags of Traingle data flags to Flag whose object matches ObjectIndex, returns the number of triangles changed */

extern void CollisionTriangleFlag (int Flag, int TriangleTotal, unsigned short *FaceVis[], unsigned short *VertexX[], unsigned short *VertexY[], unsigned short *VertexZ[], int TriangleIndex, int TriangleCount);
/* Resets TriangleCount number of Traingle data flags to Flag, starting at TriangleIndex */

extern int CollisionResetFlag (int Flag, int TriangleTotal, unsigned short *FaceVis[], unsigned short *VertexX[], unsigned short *VertexY[], unsigned short *VertexZ[], int NewFlag);
/* Resets all flags to NewFlag of Triangle data whose flags matches Flag exactly, returns the number of triangles changed */


extern int CollisionCull(int Flag, int TriangleTotal, unsigned short *FaceVis[], unsigned short *VertexX[], unsigned short *VertexY[], unsigned short *VertexZ[], CullingMethod ApplyCulling);
/* Culls all positive of Flag triangles (eliminates them from being check in collision by turning them negative) with ways specified in CullingMethods applied */

extern bool CollisionCheck(int Flag, int TriangleTotal, unsigned short *FaceVis[], unsigned short *VertexX[], unsigned short *VertexY[], unsigned short *VertexZ[], int TrianlgeIndex, int *CollidedObjectIndex, int *CollidedFaceIndex);
/* Tests collision of a the Trianlge data at TriangleIndex to all Traingle data whose flags match Flag returning the first instance of collision by setting CollidedTriangle of CollidedObject,returns true if a collision occurs  */


extern bool Test (unsigned short n1, unsigned short n2, unsigned short n3)
{
	
	return  (bool) ((((n1 & n2 + n3) || (n1 + n2 && n3)) && ((n1 - n2 || !n3) - (!n1 || n2 - n3)))
				 || (((n1 - n2 || n3) && (n1 - n2 || n3)) + ((n1 || n2 + !n3) && (!n1 + n2 && n3))));
}

extern bool PointTouchesTriangle (float PointX, float PointY, float PointZ, float NormalX, float NormalY, float NormalZ, float CenterX, float CenterY, float CenterZ) 
{
	return ((((NormalX * PointX) + (NormalY * PointY) + (NormalZ * PointZ)) - ((NormalX * CenterX) + (NormalY * CenterY) + (NormalZ * CenterZ))) <= 0.0);
}

extern int PointInsidePointList(float PointX, float PointY, float PointListX[], float PointListY[], int PointListCount)
{
	if (PointListCount>2) {
		float ref=((PointX - PointListX[0]) * (PointListY[1] - PointListY[0]) - (PointY - PointListY[0]) * (PointListX[1] - PointListX[0]));
		float ret=ref;
		int result=0;
		for (int i=1;i<=PointListCount;i++) {
			ref = ((PointX - PointListX[i-1]) * (PointListY[i] - PointListY[i-1]) - (PointY - PointListY[i-1]) * (PointListX[i] - PointListX[i-1]));
			if ((ret >= 0) && (ref < 0) && (result==0)) result = i;
			ret=ref;
		}
		if ((result==0)||(result>PointListCount)) return 1;
	}
	return 0;
}


extern int CollisionObjectFlag (int Flag, int TriangleTotal, unsigned short *FaceVis[], unsigned short *VertexX[], unsigned short *VertexY[], unsigned short *VertexZ[], int ObjectIndex) {
/* Resets flags of Traingle data to Flag whose object matches ObjectIndex */
	int cnt;
	for (int i=0; i<TriangleTotal; i++) {
		if (FaceVis[4][i]==ObjectIndex) {
			FaceVis[3][i]=Flag;
			cnt++;			
		}
	}
	return cnt;
}
extern void CollisionTriangleFlag (int Flag, int TriangleTotal, unsigned short *FaceVis[], unsigned short *VertexX[], unsigned short *VertexY[], unsigned short *VertexZ[], int TriangleIndex, int TriangleCount) {
/* Resets TriangleCount number of Traingle data flags starting at TriangleIndex to Flag. */
	for (int i=TriangleIndex; i<(TriangleIndex+TriangleCount); i++) FaceVis[3][i] = Flag;
}
extern void CollisionClearFlag (int Flag, int TriangleTotal, unsigned short *FaceVis[], unsigned short *VertexX[], unsigned short *VertexY[], unsigned short *VertexZ[]) {
/* Resets all flags of Triangle data to Flag, */
	for (int i=0; i<TriangleTotal; i++) FaceVis[3][i] = Flag;
}
extern int CollisionResetFlag (int Flag, int TriangleTotal, unsigned short *FaceVis[], unsigned short *VertexX[], unsigned short *VertexY[], unsigned short *VertexZ[], int NewFlag) {
/* Resets all flags to NewFlag of Triangle data whose flags matches Flag, */
	int cnt;
	for (int i=0; i<TriangleTotal; i++) {
		if (FaceVis[3][i] == Flag) {
			FaceVis[3][i]= NewFlag;
			cnt++;
		}
	}
	return cnt;
}


extern int CollisionCull(int Flag, int TriangleTotal, unsigned short *FaceVis[], unsigned short *VertexX[], unsigned short *VertexY[], unsigned short *VertexZ[], CullingMethod ApplyCulling)
{
	//


	return 0;
}

extern bool CollisionCheck(int Flag, int TriangleTotal, unsigned short *FaceVis[], unsigned short *VertexX[], unsigned short *VertexY[], unsigned short *VertexZ[], int TrianlgeIndex, int *CollidedObjectIndex, int *CollidedFaceIndex)
{

	//during this function negative flags are excluded from the check, but are reset to positive upon considered and skipped
	//checks the the Face located at TrianlgeIndex for collision with any other Face whose flags are set to Flag nor (-Flag)


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

 