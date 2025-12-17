#include <math.h>
#include <stdio.h> 

/*
float TriangleSetData[6][#];
TriangleFaceData dimension [][n] where n=# is triangle index
TriangleFaceData dimension [n][] where n=0 is x of the triangle normal
TriangleFaceData dimension [n][] where n=1 is y of the triangle normal
TriangleFaceData dimension [n][] where n=2 is z of the triangle normal
TriangleFaceData dimension [n][] where n=3 custom flag for segragation
TriangleFaceData dimension [n][] where n=4 a object organization index
TriangleFaceData dimension [n][] where n=5 network balance flag states
float VertexXAxisData[3][#];
VertexXAxisData dimension [][n] where n=# is triangle index
VertexXAxisData dimension [n][] where n=0 is X of the first vertex
VertexXAxisData dimension [n][] where n=1 is X of the second vertex
VertexXAxisData dimension [n][] where n=2 is X of the third vertex
float VertexYAxisData[3][#];
VertexYAxisData dimension [][n] where n=# is triangle index
VertexYAxisData dimension [n][] where n=0 is Y of the first vertex
VertexYAxisData dimension [n][] where n=1 is Y of the second vertex
VertexYAxisData dimension [n][] where n=2 is Y of the third vertex
float VertexZAxisData[3][#];
VertexZAxisData dimension [][n] where n=# is triangle index
VertexZAxisData dimension [n][] where n=0 is Z of the first vertex
VertexZAxisData dimension [n][] where n=1 is Z of the second vertex
VertexZAxisData dimension [n][] where n=2 is Z of the third vertex

 */
#define NORMAL_X 0
#define NORMAL_Y 1
#define NORMAL_Z 2

#define CUSTOM_FLAG 3
#define OBJECT_INDEX 4
#define CULLED_FLAG 5

#define VERTEX1 0
#define VERTEX2 1
#define VERTEX3 2


/*
//the higher the culling method the more inclusive the triangles are vs apply all culling in potential defaults simply
#define CullByFlagElimination 0  //if culling or flags are used, this is automatic, or else csonder using flag 0 every call
#define CullBySquareBoundary 1  //a laments term version of ByCameras, this defined rectangles for three axis that encumbant the triangle and eliminates all trainalges not found with-in the rectangles entirety, has near issues, it can be a fast elimination processes
#define CullByInside2DShape 2  //similar to squares but defined by a 2D complex shape view of the scene, points are tested per 3 axis whether or not they are inside any complex 2D shape determined by Flag, where it is not, are elminated from from the check
#define CullByProximityRange 3  //this eliminates triangles by a range factor from the center of the test triangle to it's potential rnage, all three lengths added up, as max permititer of a spherical catch for other traingles centers and lengths robustly eliminating
#define CullByClosestDistance 4  //this is a more refined version of Ranging, it attempts to determine the closest traingle to the test traingle by points, more effective with other culliing preformed first to wean down its wean it checks point for point via distance
#define CullByCustomizedCamera 5  //a very strong control to approch of culling effective in multiple call projections, applying a Up/Eye/Dir anything with in view, like a rectangle projection only customized and not so horizontal and vertical locked, repeatable
#define CullByBackfaceExclusion 6  //by defualt culling does not include backfacings, unless you include it with this flag, this will attempt to see all trinalges tested under 45 degrees of similarity as eliminated, and those facing each other as the only in determination
#define UseAllCullingMethods 7  //apply all the culling methods above that are possible with the input given in a probable best case scenario and not perfect
*/


#define CullByFlagElimination 0
#define CullBySquareBoundary 1
#define CullByInside2DShape 2
#define CullByProximityRange 3
#define CullByClosestDistance 4
#define CullByCustomizedCamera 5
#define CullByBackfaceExclusion 6
#define UseAllCullingMethods 7


struct Point {
    float X;
    float Y;
    float Z;
};

const double epsilon = 1e-9;


extern int PointTouchesTriangle(float PointX, float PointY, float PointZ, float NormalX, float NormalY, float NormalZ, float CenterX, float CenterY, float CenterZ);
/* Checks for the presence of a point possibly behind a triangle, the first three inputs are the point to test with
the triangles center removed, the next three are the triangles normal, the last three are tthe triangles center. */

extern int PointInsidePointList(float PointX, float PointY,float *PointListX, float *PointListY, int PointListCount);
/* Tests for the presence of a 2D point pointX,pointY anywhere within a 2D shape defined with a list of points pointListX,pointListY that has pointListCount number of coordinates,
returning the the unsigned percentage of maximum datatype numerical relation to percentage of total coordinates, or zero if the point does not occur within the shapes defined boundaries. */

extern float TriangleCrossSegment(float Ax1, float Ay1, float Az1, float Ax2, float Ay2, float Az2, float Ax3, float Ay3, float Az3, float Bx1, float By1, float Bz1, float Bx2, float By2, float Bz2, float Bx3, float By3, float Bz3);

extern float TriangleCrossSegmentEx(float Ax1, float Ay1, float Az1, float Ax2, float Ay2, float Az2, float Ax3, float Ay3, float Az3, float Bx1, float By1, float Bz1, float Bx2, float By2, float Bz2, float Bx3, float By3, float Bz3, float *Px0, float *Py0, float *Pz0, float *Px1, float *Py1, float *Pz1);
/* Accepts two trianlges, A, and B, by 3 verticies each, and returns the length of the overlapping segment
line formed from their collision, as well the points of the line segment if the extended version */


extern void FlagClear(int Flag, int TriangleTotal, float *FaceVis, float *VertexX, float *VertexY, float *VertexZ);
/* Resets all flags to Flag of Triangle data, */

extern int FlagObject(int Flag, int TriangleTotal, float *FaceVis, float *VertexX, float *VertexY, float *VertexZ, int ObjectIndex);
/* Resets flags of Traingle data flags to Flag whose object matches ObjectIndex, returns the number of triangles changed */

extern void FlagTriangle (int Flag, int TriangleTotal, float *FaceVis, float *VertexX, float *VertexY, float *VertexZ, int TriangleIndex, int TriangleCount);
/* Resets TriangleCount number of Traingle data flags to Flag, starting at TriangleIndex */

extern int FlagModify (int Flag, int TriangleTotal, float *FaceVis, float *VertexX, float *VertexY, float *VertexZ, int NewFlag);
/* Resets all flags to NewFlag of Triangle data whose flags matches Flag exactly, returns the number of triangles changed */

extern void ResetCulling (int Flag, int TriangleTotal, float FaceVis[], float VertexX[], float VertexY[], float VertexZ[]) ;
/* Resets all temporary culling flags of Triangle data whose flag matches Flag, if Flag is zero all culled triangles are reset */

extern int CollisionCull(int Flag, int TriangleTotal, float *FaceVis, float *VertexX, float *VertexY, float *VertexZ, int TriangleIndex, int ApplyCulling);
/* Culls triangles (eliminates them from being checked CollisionCheck) weaning by Flag first, before the CullingMethods applied, and can be sequentally called to refinement or all/partial includive in one call */

extern bool CollisionCheck(int Flag, int TriangleTotal, float *FaceVis, float *VertexX, float *VertexY, float *VertexZ, int TriangleIndex, int *CollidedObjectIndex, int *CollidedTriangleIndex);
/* Tests collision of a the Trianlge data at TriangleIndex to all Traingle data whose flags match Flag returning true if a collision occurs setting CollidedTriangle of CollidedObject*/

// ===== Helpers =====
float RoundN(float value, int n)
{
    float factor = powf(10.0f, (float)n);
    return floorf(value * factor + 0.5f) / factor;
}


bool checkbit(int bits, int num) {
	return ((bits & (1 << num))>0);
}

void setbit(int bits, int num) {
	bits |= (1 << num);
}
void clearbit(int bits, int num) {
	bits &= ~(1 << num);
}
void togglebit(int bits, int num) {
	bits ^= (1 << num);
}

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

float Least(float a, float b) { return ((a < b) ? a : b); }
float Least(float a, float b, float c) { return ((a < b) ? ((a < c) ? a : c) : ((b < c) ? b : c)); }
float Large(float a, float b) { return ((a > b) ? a : b); }
float Large(float a, float b, float c) { return ((a > b) ? ((a > c) ? a : c) : ((b > c) ? b : c)); }


Point VectorCrossProduct(Point a, Point b) {
    Point r;
    r.X = a.Y*b.Z - a.Z*b.Y;
    r.Y = a.Z*b.X - a.X*b.Z;
    r.Z = a.X*b.Y - a.Y*b.X;
    return r;
}
float Distance(float p1X,float p1Y, float p1Z, float p2X, float p2Y, float p2Z) {
    float dx = (p1X - p2X);
    float dy = (p1Y - p2Y);
    float dz = (p1Z - p2Z);
    float sumSq = dx*dx + dy*dy + dz*dz;
    return (sumSq != 0.0) ? sqrtf(sumSq) : 0;
}
float DistanceEx(Point p1, Point p2) {
    float dx = p1.X - p2.X;
    float dy = p1.Y - p2.Y;
    float dz = p1.Z - p2.Z;
    float sumSq = dx*dx + dy*dy + dz*dz;
    return (sumSq != 0.0) ? sqrtf(sumSq) : 0;
}

Point TriangleAxii(Point p1, Point p2, Point p3) {
	return MakePoint((Least(p1.X, p2.X, p3.X) + ((Large(p1.X, p2.X, p3.X) - Least(p1.X, p2.X, p3.X)) / 2)),
					 (Least(p1.Y, p2.Y, p3.Y) + ((Large(p1.Y, p2.Y, p3.Y) - Least(p1.Y, p2.Y, p3.Y)) / 2)),
			   	 	 (Least(p1.Z, p2.Z, p3.Z) + ((Large(p1.Z, p2.Z, p3.Z) - Least(p1.Z, p2.Z, p3.Z)) / 2)));
}
Point TriangleOffset(Point p1, Point p2, Point p3) {
    return MakePoint((Large(p1.X, p2.X, p3.X) - Least(p1.X, p2.X, p3.X)),
				 	 (Large(p1.Y, p2.Y, p3.Y) - Least(p1.Y, p2.Y, p3.Y)),
					 (Large(p1.Z, p2.Z, p3.Z) - Least(p1.Z, p2.Z, p3.Z)));
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

Point FacePlaneNormal(Point Up, Point Direction) {
    // Compute Right vector
    Point Right = VectorCrossProduct(Direction, Up);

    // Plane normal is perpendicular to Right and Up
    Point Normal = VectorCrossProduct(Right, Up);

    return VectorNormalize(Normal);
}


Point PlaneNormal(Point p1, Point p2, Point p3) {
    return VectorNormalize(TriangleNormal(p1, p2, p3));
}

bool PlaneBackfaceToPlane(float nx1, float ny1, float nz1,
						  float nx2, float ny2, float nz2) {
	return (!((((nx1+nx2) > -0.50f) && ((nx1+nx2) < 0.50f)) &&
			(((ny1+ny2) > -0.50f) && ((ny1+ny2) < 0.50f)) && 
			(((nz1+nz2) > -0.50f) && ((nz1+nz2) < 0.50f))));
}

bool Test (unsigned short n1, unsigned short n2, unsigned short n3)
{
	
	return  (bool) ((((n1 & n2 + n3) || (n1 + n2 && n3)) && ((n1 - n2 || !n3) - (!n1 || n2 - n3)))
				 || (((n1 - n2 || n3) && (n1 - n2 || n3)) + ((n1 || n2 + !n3) && (!n1 + n2 && n3))));
}

float Sign(float n) {
	//returns the sign of any number which is the multiplication facttr of it's negative (*-1), zero(*0) or positive (*1)
	return RoundN(((-(fabsf((n*99.99f)-1) - (n*99.99f)) - (-fabsf((n*99.99f)+1) + (n*99.99f)))* 0.5f),0);
}

/*
extern bool PointBehindPoly (float pointX, float pointY, float pointZ, float length1, float length2, float length3, float normalX, float normalY, float normalZ) 
{
	return (pointZ * length3 + length2 * pointY + length1 * pointX) - (length3 * normalZ + length1 * normalX + length2 * normalY) <= 0.0;
}*/


extern int PointTouchesTriangle(float PointX, float PointY, float PointZ, float Length1, float Length2, float Length3, float NormalX, float NormalY, float NormalZ) 
{
	return (( (Length(MakePoint(PointX, PointY, PointZ)) <=
			( ( ((Length1+Length2)/2.0f)+((Length1+Length3)/2.0f)+((Length2+Length3)/2.0f)  ) /3.0f )) &&	
			( (Sign(PointX)>=Sign(NormalX)) && (Sign(PointY)>=Sign(NormalY)) && (Sign(PointZ)>=Sign(NormalZ)) )) ) ? -1 :0;
}

extern int PointInsidePointList ( float pointX, float pointY, float polyDataX[], float polyDataY[], int polyDataCount)
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
		if ((result==0)||(result>polyDataCount)) {
			return -1;
		} else {
			return result;
		}
	}
	return 0;
}

/*
extern bool PointInsidePointList(float PointX, float PointY, float *PointListX, float *PointListY, int PointListCount)
{
    bool inside = false;

    for (int i = 0, j = PointListCount - 1; i < PointListCount; j = i++)
    {
        float xi = PointListX[i], yi = PointListY[i];
        float xj = PointListX[j], yj = PointListY[j];

        // Check if the edge crosses the horizontal ray to the right of (px, py)
        bool intersect =
            ((yi > PointY) != (yj > PointY)) &&
            (PointX < (xj - xi) * (PointY - yi) / (yj - yi + 1e-15) + xi);

        if (intersect)
            inside = !inside;
    }

    return inside;
}
*/

extern void FlagClear (int Flag, int TriangleTotal, float FaceVis[], float VertexX[], float VertexY[], float VertexZ[]) {
/* Resets all flags of Triangle data to Flag, */
	for (int i=0; i<TriangleTotal; i++) {
		FaceVis[(i*6)+CUSTOM_FLAG] = (float)Flag;
	}
}

extern int FlagObject (int Flag, int TriangleTotal, float FaceVis[], float VertexX[], float VertexY[], float VertexZ[], int ObjectIndex) {
/* Resets flags of Traingle data to Flag whose object matches ObjectIndex */
	int cnt=0;
	for (int i=0; i<TriangleTotal; i++) {
		if (FaceVis[(i*6)+OBJECT_INDEX]==(float)ObjectIndex) {
			FaceVis[(i*6)+CUSTOM_FLAG]=(float)Flag;
			cnt++;			
		}
	}
	return cnt;
}
extern void FlagTriangle (int Flag, int TriangleTotal, float FaceVis[], float VertexX[], float VertexY[], float VertexZ[], int TriangleIndex, int TriangleCount) {
/* Resets TriangleCount number of Traingle data flags starting at TriangleIndex to Flag. */
	for (int i=(TriangleIndex+1); i<=(TriangleIndex+TriangleCount); i++) {FaceVis[(i*6)+CUSTOM_FLAG] = (float)Flag;}
}

extern int FlagModify (int Flag, int TriangleTotal, float FaceVis[], float VertexX[], float VertexY[], float VertexZ[], int NewFlag) {
/* Resets all flags to NewFlag of Triangle data whose flags matches Flag, */
	int cnt=0;
	for (int i=0; i<TriangleTotal; i++) {
		if (fabsf(FaceVis[(i*6)+CUSTOM_FLAG]) == Flag) {
			FaceVis[(i*6)+CUSTOM_FLAG]= (float)NewFlag;
			cnt++;
		}
	}
	return cnt;
}

extern void ResetCulling (int Flag, int TriangleTotal, float FaceVis[], float VertexX[], float VertexY[], float VertexZ[]) {
/* Resets all temporary culling flags of Triangle data whose flag matches Flag, if Flag is zero all culled triangles are reset */
	if (Flag==0) {
		for (int i=0; i<TriangleTotal; i++) {
			FaceVis[(i*6)+CULLED_FLAG] = fabsf((float)Flag);
		}
	} else {
		for (int i=0; i<TriangleTotal; i++) {
			if (fabsf(FaceVis[(i*6)+CUSTOM_FLAG])==(float)Flag) {
				FaceVis[(i*6)+CULLED_FLAG] = fabsf((float)Flag);
			}
		}	
	}
}

bool SkipCollisionCheck(int Flag, float FaceVis[], int TriangleIndex, int i) {
	//a very basic simply avoidance check for the collision function to ignore by flag and the triangle checked
	return ((FaceVis[(i*6)+CUSTOM_FLAG]==Flag)&&(FaceVis[(i*6)+OBJECT_INDEX]!=FaceVis[(TriangleIndex*6)+OBJECT_INDEX])&&(FaceVis[(i*6)+CULLED_FLAG]==Flag));
}

bool SkipCullingCheck(int Flag, float FaceVis[], int TriangleIndex, int i, int ApplyCulling) {
	bool ret=true;

	//if (fabsf(FaceVis[(i*6)+5])!=Flag) FaceVis[(i*6)+5]=(float)Flag;//init the flag's modifier
	ret =  SkipCollisionCheck(Flag,FaceVis,TriangleIndex,i);
	if  (!ret) {
		if (FaceVis[(6*i)+CULLED_FLAG]>0) { //ignore those already culled

			if (checkbit(ApplyCulling,CullByBackfaceExclusion)||checkbit(ApplyCulling,UseAllCullingMethods)) {
				//backfacing should maybe be done at this level similarly low level as the flagging
				if (PlaneBackfaceToPlane(FaceVis[(6*TriangleIndex)+NORMAL_X],FaceVis[(6*TriangleIndex)+NORMAL_Y],FaceVis[(6*TriangleIndex)+NORMAL_Z],
					FaceVis[(6*i)+NORMAL_X],FaceVis[(6*i)+NORMAL_Y],FaceVis[(6*i)+NORMAL_Z])) FaceVis[(6*i)+CULLED_FLAG] = -FaceVis[(6*i)+CULLED_FLAG]; //flag it off

			}
		}
	}
	return ret;
}
void Get2Dfrom3Dpoints(float vertexList[],int start, int count,float *pointListOut, int *size) {
	
	float *temp=new float[count*3];

	int i=0;
	while (i<(count*3))  {
		temp[i+VERTEX1] = vertexList[(3*(start+i))+VERTEX1];
		temp[i+VERTEX2] = vertexList[(3*(start+i))+VERTEX1];
		temp[i+VERTEX3] = vertexList[(3*(start+i))+VERTEX1];
		i=i+3;
		
	}

	*size = i;
	delete[] pointListOut;
	*pointListOut = *temp;

}

extern int CollisionCull(int Flag, int TriangleTotal, float FaceVis[], float VertexX[], float VertexY[], float VertexZ[], int TriangleIndex, int ApplyCulling)
{
	int start=0, sstop=0, count=0,i=0, culled=0, inc=0;
	float minX=0, maxX=0, minY=0, maxY=0, minZ=0, maxZ=0;
	float test=0, dist=0;
	float prox=0;
	int *size=0;
	float *pointListX=new float[];
	float *pointListY=new float[];
	float *pointListZ=new float[];

	Point p1,p2,p3,c1,c2;

	while (i<TriangleTotal) {

		if (SkipCullingCheck(Flag,FaceVis,TriangleIndex,i,ApplyCulling)) {
			start = i;
			sstop = i;

			while ( SkipCullingCheck(Flag,FaceVis,TriangleIndex,sstop,ApplyCulling)) {
				sstop++;
				if (sstop>=TriangleTotal) break;
			}

			if (sstop>start) {
				sstop = (sstop-1);
				count = (sstop - (start-1));
				if ( ((start + (count -1)) <= TriangleTotal) && (start <=sstop)) {

					/*
					#define CullByFlagElimination 0
					#define CullBySquareBoundary 1
					#define CullByInside2DShape 2
					#define CullByProximityRange 3
					#define CullByClosestDistance 4
					#define CullByCustomizedCamera 5
					#define CullByBackfaceExclusion 6
					#define UseAllCullingMethods 7
					*/

					//############################ PREP

					if (checkbit(ApplyCulling,CullBySquareBoundary)||checkbit(ApplyCulling,UseAllCullingMethods)) {
						//gather a min and max with anything not culled
						//abvoe zero is not culled
						for (int j = 0; j<3; j++) {
							//always include the test triangle
							if ((VertexX[(3*TriangleIndex)+j]>maxX)||(maxX==0)) maxX=VertexX[(3*TriangleIndex)+j];
							if ((VertexX[(3*TriangleIndex)+j]>maxY)||(maxY==0)) maxY=VertexY[(3*TriangleIndex)+j];
							if ((VertexX[(3*TriangleIndex)+j]>maxZ)||(maxZ==0)) maxZ=VertexZ[(3*TriangleIndex)+j];
							
							if ((VertexX[(3*TriangleIndex)+j]<minX)||(minX==0)) minX=VertexX[(3*TriangleIndex)+j];
							if ((VertexX[(3*TriangleIndex)+j]<minY)||(minY==0)) minY=VertexY[(3*TriangleIndex)+j];
							if ((VertexX[(3*TriangleIndex)+j]<minZ)||(minZ==0)) minZ=VertexZ[(3*TriangleIndex)+j];
						}
						for (int t = start; t < count; t+=3)
						{//object iteration
							for (int j = 0; j<3; j++) {
								if (FaceVis[(6*t)+CULLED_FLAG]>0) { //ignore those already culled
									if ((VertexX[(3*t)+j]>maxX)||(maxX==0)) maxX=VertexX[(3*t)+j];
									if ((VertexX[(3*t)+j]>maxY)||(maxY==0)) maxY=VertexY[(3*t)+j];
									if ((VertexX[(3*t)+j]>maxZ)||(maxZ==0)) maxZ=VertexZ[(3*t)+j];
									
									if ((VertexX[(3*t)+j]<minX)||(minX==0)) minX=VertexX[(3*t)+j];
									if ((VertexX[(3*t)+j]<minY)||(minY==0)) minY=VertexY[(3*t)+j];
									if ((VertexX[(3*t)+j]<minZ)||(minZ==0)) minZ=VertexZ[(3*t)+j];
								}
							}
						}
						
					}
					if (checkbit(ApplyCulling,CullByInside2DShape)||checkbit(ApplyCulling,UseAllCullingMethods)) {

						Get2Dfrom3Dpoints(VertexX, start, count, pointListX, size);
						Get2Dfrom3Dpoints(VertexY, start, count, pointListY, size);
						Get2Dfrom3Dpoints(VertexZ, start, count, pointListZ, size);

					}

					if (checkbit(ApplyCulling,CullByProximityRange)||checkbit(ApplyCulling,UseAllCullingMethods)) {
						//should go before CullByClosestDistance, if both are in one call
						//due to the exacting preformance in large traingle mapping
						//this is far more robust then CullByClosestDistance

						p1=MakePoint(VertexX[(3*TriangleIndex)+VERTEX1],VertexY[(3*TriangleIndex)+VERTEX1],VertexZ[(3*TriangleIndex)+VERTEX1]);
						p2=MakePoint(VertexX[(3*TriangleIndex)+VERTEX2],VertexY[(3*TriangleIndex)+VERTEX2],VertexZ[(3*TriangleIndex)+VERTEX2]);
						p3=MakePoint(VertexX[(3*TriangleIndex)+VERTEX3],VertexY[(3*TriangleIndex)+VERTEX3],VertexZ[(3*TriangleIndex)+VERTEX3]);

						c1=TriangleAxii(p1,p2,p3);

						prox = DistanceEx(p1,p2) + DistanceEx(p2,p3) + DistanceEx(p3,p1);

					}
					if (checkbit(ApplyCulling,CullByClosestDistance)||checkbit(ApplyCulling,UseAllCullingMethods)) {
						//gather the shortest distance of any point to any the test triangle
						//abvoe zero is not culled
						for (int j = 0; j<3; j++) {
							//always include the test triangle
							for (int t = start; t < count; t+=3)
							{//object iteration
								for (int jj = 0; jj<3; jj++) {
									if (FaceVis[(6*t)+CULLED_FLAG]>0) { //ignore those already culled
										test=Distance(VertexX[(3*TriangleIndex)+j],VertexY[(3*TriangleIndex)+j],VertexZ[(3*TriangleIndex)+j],
											VertexX[(3*t)+jj],VertexY[(3*t)+jj],VertexZ[(3*t)+jj]);
										if (test<dist) dist = test;
									}
								}
							}
						}
					}

					//############################ FLAGS

					for (int t = start; t < count; t++)
					{//object iteration

						if (checkbit(ApplyCulling,CullBySquareBoundary)||checkbit(ApplyCulling,UseAllCullingMethods)) {
							//this one is going to fastly cut anthing done below down
							//quicker and may render CullByInside2DShape obsolete due to that
							if (FaceVis[(6*t)+CULLED_FLAG]>0) {//still part of the collisioncheck
								for (int j = 0; j<3; j++) {
									if (((VertexX[(3*t)+j]<minX)&&(VertexX[(3*t)+j]>maxX))&&
										((VertexX[(3*t)+j]<minX)&&(VertexX[(3*t)+j]>maxX)) &&
										((VertexX[(3*t)+j]<minX)&&(VertexX[(3*t)+j]>maxX))) {
										inc++; //count the number outside the bound rect

									}
								}
								if (inc==3) FaceVis[(6*t)+CULLED_FLAG] = -FaceVis[(6*t)+CULLED_FLAG]; //flag it off
								inc=0;
							}

						}
						if (checkbit(ApplyCulling,CullByInside2DShape)||checkbit(ApplyCulling,UseAllCullingMethods)) {

							if (FaceVis[(6*t)+CULLED_FLAG]>0) {//still part of the collisioncheck
								inc=0;
								for (int j=0;j<3;j++) {
									if (PointInsidePointList(VertexX[(3*t)+j],VertexY[(3*t)+j],pointListX,pointListY,count)) {
										inc++;
										if (inc>2) break;
										if (PointInsidePointList(VertexY[(3*t)+j],VertexZ[(3*t)+j],pointListY,pointListZ,count)) {
											inc++;
											if (inc>2) break;
											if (PointInsidePointList(VertexZ[(3*t)+j],VertexX[(3*t)+j],pointListZ,pointListX,count)) {
												inc++;
												if (inc>2) break;
											}
										}
									}
								}
								if (inc<3) FaceVis[(6*t)+CULLED_FLAG] = -FaceVis[(6*t)+CULLED_FLAG]; //flag it off
							}

						}

						if (checkbit(ApplyCulling,CullByProximityRange)||checkbit(ApplyCulling,UseAllCullingMethods)) {
							//should go before CullByClosestDistance, if both are in one call
							//due to the exacting preformance in large traingle mapping
							//this is far more robust then CullByClosestDistance

							if (FaceVis[(6*t)+CULLED_FLAG]>0) {//still part of the collisioncheck
								p1=MakePoint(VertexX[(3*t)+VERTEX1],VertexY[(3*t)+VERTEX1],VertexZ[(3*t)+VERTEX1]);
								p2=MakePoint(VertexX[(3*t)+VERTEX2],VertexY[(3*t)+VERTEX2],VertexZ[(3*t)+VERTEX2]);
								p3=MakePoint(VertexX[(3*t)+VERTEX3],VertexY[(3*t)+VERTEX3],VertexZ[(3*t)+VERTEX3]);

								c2=TriangleAxii(p1,p2,p3);

								if (DistanceEx(c1,c2)>prox)  FaceVis[(6*t)+CULLED_FLAG] = -FaceVis[(6*t)+CULLED_FLAG];
							}


						}
						if (checkbit(ApplyCulling,CullByClosestDistance)||checkbit(ApplyCulling,UseAllCullingMethods)) {
							inc=9;
							for (int j = 0; j<3; j++) {								
								for (int jj = 0; jj<3; jj++) {
									if (FaceVis[(6*t)+CULLED_FLAG]>0) { //ignore those already culled
										test=Distance(VertexX[(3*TriangleIndex)+j],VertexY[(3*TriangleIndex)+j],VertexZ[(3*TriangleIndex)+j],
												VertexX[(3*t)+jj],VertexY[(3*t)+jj],VertexZ[(3*t)+jj]);
										if (test>dist) {
											break;
										} else {
											inc--;
										}
									}	
								}
								if ((inc!=6)&&(inc!=3)) break;

							}
							if (inc!=0) FaceVis[(6*t)+CULLED_FLAG] = -FaceVis[(6*t)+CULLED_FLAG]; //flag it off
							inc=0;
						}
					}
				}
				i = sstop;
			}
		}
		i++;	
	}	

	return culled;

}

extern bool CollisionCheck(int Flag, int TriangleTotal, float FaceVis[], float VertexX[], float VertexY[], float VertexZ[], int TriangleIndex, int *CollidedObjectIndex, int *CollidedTriangleIndex)
{
	int i=0;
	float lx=0,ly=0,lz=0,nx=0,ny=0,nz=0,cx=0,cy=0,cz=0;
	Point p1,p2,p3;

	while  (i<TriangleTotal) {		

		if (!SkipCollisionCheck(Flag,FaceVis,TriangleIndex,i)) {
			//the flag is equal to the one we want, and the objectindex is not the same as the triangle we are checking			

			p1=MakePoint(VertexX[(3*i)+VERTEX1],VertexY[(3*i)+VERTEX1],VertexZ[(3*i)+VERTEX1]);
			p2=MakePoint(VertexX[(3*i)+VERTEX2],VertexY[(3*i)+VERTEX2],VertexZ[(3*i)+VERTEX2]);
			p3=MakePoint(VertexX[(3*i)+VERTEX3],VertexY[(3*i)+VERTEX3],VertexZ[(3*i)+VERTEX3]);

			lx=DistanceEx(p1,p2);
			ly=DistanceEx(p2,p3);
			lz=DistanceEx(p3,p1);

			cx=Least(p1.X, p2.X, p3.X);
			cy=Least(p1.Y, p2.Y, p3.Y);
			cz=Least(p1.Z, p2.Z, p3.Z);
			cx=(cx + ((Large(p1.X, p2.X, p3.X) - cx) / 2));
			cy=(cy + ((Large(p1.Y, p2.Y, p3.Y) - cy) / 2));
			cz=(cz + ((Large(p1.Z, p2.Z, p3.Z) - cz) / 2));

			nx=FaceVis[(6*i)+NORMAL_X];
			ny=FaceVis[(6*i)+NORMAL_Y];
			nz=FaceVis[(6*i)+NORMAL_Z];

			for (int j=0;j<3;j++) {

				if (PointTouchesTriangle(VertexX[(3*TriangleIndex)+j]-cx,VertexY[(3*TriangleIndex)+j]-cy,VertexZ[(3*TriangleIndex)+j]-cz,lx,ly,lz,nx,ny,nz)) {
					*CollidedObjectIndex=(int)FaceVis[(i*6)+OBJECT_INDEX];
					*CollidedTriangleIndex=i;
					return true;
				}	
			}
			
		}
		i=i+3;				
	}
	return false;
	
}



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


float PlaneAngleToPlane(float nx1, float ny1, float nz1,
						float nx2, float ny2, float nz2)
{
    // Dot product
    float dot = nx1*nx2 + ny1*ny2 + nz1*nz2;

    // Magnitudes
    float mag1 = sqrtf(nx1*nx1 + ny1*ny1 + nz1*nz1);
    float mag2 = sqrtf(nx2*nx2 + ny2*ny2 + nz2*nz2);

    if (mag1 == 0 || mag2 == 0) return 0.0f; // avoid divide by zero

    // Cosine of angle
    float cosTheta = dot / (mag1 * mag2);

    // Clamp to [-1, 1] to avoid floating point errors
    if (cosTheta > 1.0f) cosTheta = 1.0f;
    if (cosTheta < -1.0f) cosTheta = -1.0f;

    // Convert to degrees
    float angle = acosf(cosTheta) * 180.0f / 3.14159265f;

    return angle;
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
        l1 = DistanceEx(t1p1,t1p2) + DistanceEx(t1p2,t1p3) + DistanceEx(t1p3,t1p1);
        l2 = DistanceEx(t2p1,t2p2) + DistanceEx(t2p2,t2p3) + DistanceEx(t2p3,t2p1);
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

			return DistanceEx(MakePoint(*Px0,*Py0,*Pz0),MakePoint(*Px1,*Py1,*Pz1));
		}
	}

}

 