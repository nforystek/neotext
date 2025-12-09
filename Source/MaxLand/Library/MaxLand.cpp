#include <math.h>
#include <stdio.h> 

//
//Due to some code loss, this DLL will not function as the MaxLandLib.dll used in the gaming collision from Neotext
//I am merely trying to recouperate a DLL as functional as with the same performance but it is taking a long time.
//
extern bool Test (unsigned short n1, unsigned short n2, unsigned short n3);
/* Accepts inputs n1, n2, and n3 which can be any combination of return values from PointInPoly(X,Y), and TriTriSegment()
The idea here is that using this function allows 3D collision checks to only have to traverse two 2D point lists, and with
the precise point at which a point calls inside of a list, the TriTriSegment makes up for not having to travese the third
axis view of a 2D turned 3D.  For instance, checking a 2D view, needs three views, like Top, Left and Front and that will
cover Right, Back and Bottom in 2D naturally.  Test accepts input from two PontInPoly(X,Y) views like Top and Front and the
return values placed against the return value from TriTriSegment called when and if only if, PointInPoly(X,Y) on the top and
forn both return values positive for a collision, the TriTriSegment further validates putting all three returns to Test()*/

extern bool PointBehindPoly (float a1, float a2, float a3, float a4, float a5, float a6, float a7, float a8, float a9);
/* Checks for the presence of a point behind a triangle, the first three inputs are the length of the triangles sides, the next three are the triangles normal, the last three are the point to test with the triangles center removed. */


extern short PointInPoly(float pX, float pY,float *polyX, float *polyY, short polyN);
/* Tests for the presence of a 2D point pX,pY anywhere within a 2D shape defined with a list of points polyX,polyY that has polyN number of coordinates, returning the the unsigned percentage of maximum datatype numerical relation to percentage of total coordinates, or zero if the point does not occur within the shapes defined boundaries. */


/*
This is going to be left out, it was not able called from VB6 and I don't recall the potential reason, nevertheless, new functions prefixed with TriTriSegment will be the consideration
to replacing it no loss of original in actual use and return, but working from VB6 as well too when called.  As well renamed, to reflect better it was not similar to Thomas Moller's
tri_tri_intersect(). TriTriSegmentEncoded() is how it preformed argumetns and return while used in Collision() and not called via VB6 in my utilization of this Collision based DLL.

extern short tri_tri_intersect (unsigned short v0_0, unsigned short v0_1, unsigned short v0_2, unsigned short v1_0, unsigned short v1_1, unsigned short v1_2, unsigned short v2_0, unsigned short v2_1, unsigned short v2_2, unsigned short u0_0, unsigned short u0_1, unsigned short u0_2, unsigned short u1_0, unsigned short u1_1, unsigned short u1_2, unsigned short u2_0, unsigned short u2_1, unsigned short u2_2);
// Accepts two triangle inputs in hyperbolic paraboloid collision form and returns with in the unsiged whole the percentage of each others distance to plane as one value.  **NOTE Assumes the parameter input as triangles are TRUE for collision with one another.
*/

extern int Culling (int visType, int lngFaceCount, unsigned short *sngCamera[], unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], unsigned short *sngScreenX[], unsigned short *sngScreenY[], unsigned short *sngScreenZ[], unsigned short *sngZBuffer[]);
/* Culling function with three expirimental ways to cull, defined by visType, 0 to 2, returns the difference of input triangles. lngFaceCount, sngCamera[3 x 3], sngFaceVis[6 x lngFaceCount], sngVertexX[3 x lngFaceCount]..Y..Z, sngScreenX[3 x lngFaceCount]..Y..Z, sngZBuffer[4 x lngFaceCount].  The camera is defined by position [0,0]=X, [0,1]=Y, [0,2]=Z, direction [1,0]=X, [1,1]=Y, [1,2]=Z, and upvector [2,0]=X, [2,1]=Y, [2,2]=Z.  sngFaceVis should be initialized to zero, and sngVertex arrays are 3D coordinate equivelent to sngScreen with a screenZ buffer, and Zbuffer for the verticies. */

extern bool Collision (int visType, int lngFaceCount, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngFaceNum, int *lngCollidedBrush, int *lngCollidedFace);
/* Tests collision of a lngFaceNum against a number of visible faces, lngFaceCount, whose sngFaceVis has been defined with visType as culled with the Forystek function, and returns whether or not a collision occurs also populating the lngCollidedBrush and lngCollidedFace indicating the exact object number (brush) and face number (triangle) that has the collision impact. */

extern void EncodeTriangle(float Ax, float Ay, float Az,float Bx, float By, float Bz,float Cx, float Cy, float Cz,float* CenterX, float* CenterY, float* CenterZ, float* Nx, float* Ny, float* Nz,float* L1, float* L2, float* L3);

extern float TriTriSegmentEncoded(float CxA, float CyA, float CzA,    float NxA, float NyA, float NzA, float L1A, float L2A, float L3A, float CxB, float CyB, float CzB, float NxB, float NyB, float NzB, float L1B, float L2B, float L3B, float* Px0, float* Py0, float* Pz0, float* Px1, float* Py1, float* Pz1);
extern float TriTriSegmentEncodedLen(float CxA, float CyA, float CzA, float NxA, float NyA, float NzA, float L1A, float L2A, float L3A, float CxB, float CyB, float CzB, float NxB, float NyB, float NzB, float L1B, float L2B, float L3B);

extern float TriTriSegmentFast(float Ax1, float Ay1, float Az1, float Ax2, float Ay2, float Az2, float Ax3, float Ay3, float Az3, float Bx1, float By1, float Bz1, float Bx2, float By2, float Bz2, float Bx3, float By3, float Bz3, float* Px0, float* Py0, float* Pz0, float* Px1, float* Py1, float* Pz1);
extern float TriTriSegmentFastLen(float Ax1, float Ay1, float Az1, float Ax2, float Ay2, float Az2, float Ax3, float Ay3, float Az3, float Bx1, float By1, float Bz1, float Bx2, float By2, float Bz2, float Bx3, float By3, float Bz3);



extern bool Test (unsigned short n1, unsigned short n2, unsigned short n3)
{
	
	return  (bool) ((((n1 & n2 + n3) || (n1 + n2 && n3)) && ((n1 - n2 || !n3) - (!n1 || n2 - n3)))
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
			ref = ((pointX - polyDataX[i-1]) * (polyDataY[i] - polyDataY[i-1]) - (pointY - polyDataY[i-1]) * (polyDataX[i] - polyDataX[i-1]));
			if ((ret >= 0) && (ref < 0) && (result==0)) result = i;
			ret=ref;
		}
		if ((result==0)||(result>polyDataCount)) return 1;//todo: this is suppose to return a decimal percent
		//of the total polygon points where in is found inside, but would be positive for true, to VB6 calls
	}
	return 0;
}

extern int Culling /* was Forystek */ (int visType, int lngFaceCount, unsigned short *sngCamera[], unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], unsigned short *sngScreenX[], unsigned short *sngScreenY[], unsigned short *sngScreenZ[], unsigned short *sngZBuffer[])
{
	//todo, any awesomer culling that can be done which only applies if falls in a options flag manor defined by the user so multiple calls can paint the canvas of 
	//from single player with their bullets to online load balancing in systems processing potentially quicker when objects and their counter parts hit the scenery.
	//this function crashed at vistype>3 and it also didn't only flag the applied, it reset every flag on every call so collective map flagging was not poossible.
	return 0;
}
extern bool Collision (int visType, int lngFaceCount, unsigned short *sngFaceVis[], unsigned short *sngVertexX[], unsigned short *sngVertexY[], unsigned short *sngVertexZ[], int lngFaceNum, int *lngCollidedBrush, int *lngCollidedFace)
{
	//todo recreate as best as possible the collision functionality that the original DLL preformed quite nicely, and hope to improve of course from there on.
	return 0;
}



// Exported function: encode a triangle into 9 scalars
extern void EncodeTriangle(
    float Ax, float Ay, float Az,
    float Bx, float By, float Bz,
    float Cx, float Cy, float Cz,
    float* CenterX, float* CenterY, float* CenterZ,
    float* Nx, float* Ny, float* Nz,
    float* L1, float* L2, float* L3)
{
    // Center = centroid
    *CenterX = (Ax + Bx + Cx) / 3.0f;
    *CenterY = (Ay + By + Cy) / 3.0f;
    *CenterZ = (Az + Bz + Cz) / 3.0f;

    // Normal = (B-A) × (C-A)
    float ABx = Bx - Ax, ABy = By - Ay, ABz = Bz - Az;
    float ACx = Cx - Ax, ACy = Cy - Ay, ACz = Cz - Az;

    *Nx = ABy * ACz - ABz * ACy;
    *Ny = ABz * ACx - ABx * ACz;
    *Nz = ABx * ACy - ABy * ACx;

    // Build orthogonal basis U,V in plane
    float Ux = ABx, Uy = ABy, Uz = ABz;
    float lenU = sqrtf(Ux*Ux + Uy*Uy + Uz*Uz);
    if (lenU > 0.0f) { Ux /= lenU; Uy /= lenU; Uz /= lenU; }

    float Vx = (*Ny) * Uz - (*Nz) * Uy;
    float Vy = (*Nz) * Ux - (*Nx) * Uz;
    float Vz = (*Nx) * Uy - (*Ny) * Ux;
    float lenV = sqrtf(Vx*Vx + Vy*Vy + Vz*Vz);
    if (lenV > 0.0f) { Vx /= lenV; Vy /= lenV; Vz /= lenV; }

    // Project vertices onto U,V and compute extents
    float maxU = 0.0f, maxV = 0.0f, maxR = 0.0f;
    float VX[3] = {Ax, Bx, Cx};
    float VY[3] = {Ay, By, Cy};
    float VZ[3] = {Az, Bz, Cz};

    for (int i = 0; i < 3; i++) {
        float dx = VX[i] - *CenterX;
        float dy = VY[i] - *CenterY;
        float dz = VZ[i] - *CenterZ;

        float projU = dx*Ux + dy*Uy + dz*Uz;
        float projV = dx*Vx + dy*Vy + dz*Vz;
        float r = sqrtf(dx*dx + dy*dy + dz*dz);

        if (fabsf(projU) > maxU) maxU = fabsf(projU);
        if (fabsf(projV) > maxV) maxV = fabsf(projV);
        if (r > maxR) maxR = r;
    }

    // Store lengths
    *L1 = maxU;
    *L2 = maxV;
    *L3 = maxR;
}





extern float TriTriSegmentEncoded(
    // Triangle A
    float CxA, float CyA, float CzA,
    float NxA, float NyA, float NzA,
    float L1A, float L2A, float L3A,
    // Triangle B
    float CxB, float CyB, float CzB,
    float NxB, float NyB, float NzB,
    float L1B, float L2B, float L3B,
    // Outputs
    float* Px0, float* Py0, float* Pz0,
    float* Px1, float* Py1, float* Pz1)
{
    // --- 1) Plane constants ---
    float dA = NxA * CxA + NyA * CyA + NzA * CzA;
    float dB = NxB * CxB + NyB * CyB + NzB * CzB;

    // --- 2) Line of intersection: Q = nA × nB ---
    float Qx = NyA * NzB - NzA * NyB;
    float Qy = NzA * NxB - NxA * NzB;
    float Qz = NxA * NyB - NyA * NxB;
    float denom = Qx*Qx + Qy*Qy + Qz*Qz;
    if (denom == 0.0f) {
        *Px0 = *Py0 = *Pz0 = 0.0f;
        *Px1 = *Py1 = *Pz1 = 0.0f;
        return 0.0f; // parallel planes
    }

    // --- 3) Point on line: P = ((dA nB - dB nA) × Q) / |Q|² ---
    float tmpx = (dA * NxB - dB * NxA);
    float tmpy = (dA * NyB - dB * NyA);
    float tmpz = (dA * NzB - dB * NzA);

    float Px = tmpy * Qz - tmpz * Qy;
    float Py = tmpz * Qx - tmpx * Qz;
    float Pz = tmpx * Qy - tmpy * Qx;
    Px /= denom; Py /= denom; Pz /= denom;

    // --- 4) Fast clipping extents using L3 (bounding radius) ---
    float tAmin = -L3A, tAmax = L3A;
    float tBmin = -L3B, tBmax = L3B;

    float tmin = (tAmin > tBmin) ? tAmin : tBmin;
    float tmax = (tAmax < tBmax) ? tAmax : tBmax;

    float length = tmax - tmin;
    if (length <= 0.0f) {
        *Px0 = *Py0 = *Pz0 = 0.0f;
        *Px1 = *Py1 = *Pz1 = 0.0f;
        return -fabsf(length); // no overlap
    }

    // --- 5) Segment endpoints ---
    *Px0 = Px + tmin * Qx;
    *Py0 = Py + tmin * Qy;
    *Pz0 = Pz + tmin * Qz;
    *Px1 = Px + tmax * Qx;
    *Py1 = Py + tmax * Qy;
    *Pz1 = Pz + tmax * Qz;

    return length;
}

extern float TriTriSegmentEncodedLen(
    float CxA, float CyA, float CzA,
    float NxA, float NyA, float NzA,
    float L1A, float L2A, float L3A,
    float CxB, float CyB, float CzB,
    float NxB, float NyB, float NzB,
    float L1B, float L2B, float L3B)
{
    // Plane constants
    float dA = NxA * CxA + NyA * CyA + NzA * CzA;
    float dB = NxB * CxB + NyB * CyB + NzB * CzB;

    // Line direction Q = nA × nB
    float qx = NyA * NzB - NzA * NyB;
    float qy = NzA * NxB - NxA * NzB;
    float qz = NxA * NyB - NyA * NxB;

    // Point P = ((dA nB - dB nA) × Q) / |Q|²
    float tmpx = (dA * NxB - dB * NxA);
    float tmpy = (dA * NyB - dB * NyA);
    float tmpz = (dA * NzB - dB * NzA);

    float px = tmpy * qz - tmpz * qy;
    float py = tmpz * qx - tmpx * qz;
    float pz = tmpx * qy - tmpy * qx;

    float denom = qx*qx + qy*qy + qz*qz;
    if (denom == 0.0f) return 0.0f; // parallel planes

    px /= denom; py /= denom; pz /= denom;

    // Use bounding radius (L3) for extents
    float tAmin = -L3A, tAmax = L3A;
    float tBmin = -L3B, tBmax = L3B;

    // Overlap interval
    float tmin = (tAmin > tBmin) ? tAmin : tBmin;
    float tmax = (tAmax < tBmax) ? tAmax : tBmax;

    // Return length (positive if overlap, negative if not)
    float length = tmax - tmin;
    if (length <= 0.0f) return -fabsf(length);
    return fabsf(length);
}



extern float TriTriSegmentFast(
    // Triangle A
    float Ax1, float Ay1, float Az1,
    float Ax2, float Ay2, float Az2,
    float Ax3, float Ay3, float Az3,
    // Triangle B
    float Bx1, float By1, float Bz1,
    float Bx2, float By2, float Bz2,
    float Bx3, float By3, float Bz3,
    // Outputs
    float* Px0, float* Py0, float* Pz0,
    float* Px1, float* Py1, float* Pz1)
{
    // Normals A and B
    float nAx = (Ay2 - Ay1) * (Az3 - Az1) - (Az2 - Az1) * (Ay3 - Ay1);
    float nAy = (Az2 - Az1) * (Ax3 - Ax1) - (Ax2 - Ax1) * (Az3 - Az1);
    float nAz = (Ax2 - Ax1) * (Ay3 - Ay1) - (Ay2 - Ay1) * (Ax3 - Ax1);

    float nBx = (By2 - By1) * (Bz3 - Bz1) - (Bz2 - Bz1) * (By3 - By1);
    float nBy = (Bz2 - Bz1) * (Bx3 - Bx1) - (Bx2 - Bx1) * (Bz3 - Bz1);
    float nBz = (Bx2 - Bx1) * (By3 - By1) - (By2 - By1) * (Bx3 - Bx1);

    // Plane constants
    float dA = nAx * Ax1 + nAy * Ay1 + nAz * Az1;
    float dB = nBx * Bx1 + nBy * By1 + nBz * Bz1;

    // Line direction Q = nA × nB
    float Qx = nAy * nBz - nAz * nBy;
    float Qy = nAz * nBx - nAx * nBz;
    float Qz = nAx * nBy - nAy * nBx;

    // Point P = ((dA nB - dB nA) × Q) / |Q|²
    float Px = (dA * nBx - dB * nAx);
    float Py = (dA * nBy - dB * nAy);
    float Pz = (dA * nBz - dB * nAz);

    float tmpx = Py * Qz - Pz * Qy;
    float tmpy = Pz * Qx - Px * Qz;
    float tmpz = Px * Qy - Py * Qx;

    float denom = Qx * Qx + Qy * Qy + Qz * Qz;
    if (denom == 0.0) {
        // Parallel planes (degenerate for segment purposes)
        if (Px0) *Px0 = 0.0; if (Py0) *Py0 = 0.0; if (Pz0) *Pz0 = 0.0;
        if (Px1) *Px1 = 0.0; if (Py1) *Py1 = 0.0; if (Pz1) *Pz1 = 0.0;
        return 0.0;
    }

    Px = tmpx / denom;
    Py = tmpy / denom;
    Pz = tmpz / denom;

    // Project vertices of A to get its t-interval
    float tAmin = ((Ax1 - Px) * Qx + (Ay1 - Py) * Qy + (Az1 - Pz) * Qz) / denom;
    float tAmax = tAmin;

    float t = ((Ax2 - Px) * Qx + (Ay2 - Py) * Qy + (Az2 - Pz) * Qz) / denom;
    if (t < tAmin) tAmin = t; else if (t > tAmax) tAmax = t;

    t = ((Ax3 - Px) * Qx + (Ay3 - Py) * Qy + (Az3 - Pz) * Qz) / denom;
    if (t < tAmin) tAmin = t; else if (t > tAmax) tAmax = t;

    // Project vertices of B to get its t-interval
    float tBmin = ((Bx1 - Px) * Qx + (By1 - Py) * Qy + (Bz1 - Pz) * Qz) / denom;
    float tBmax = tBmin;

    t = ((Bx2 - Px) * Qx + (By2 - Py) * Qy + (Bz2 - Pz) * Qz) / denom;
    if (t < tBmin) tBmin = t; else if (t > tBmax) tBmax = t;

    t = ((Bx3 - Px) * Qx + (By3 - Py) * Qy + (Bz3 - Pz) * Qz) / denom;
    if (t < tBmin) tBmin = t; else if (t > tBmax) tBmax = t;

    // Overlap interval
    float tmin = (tAmin > tBmin) ? tAmin : tBmin;
    float tmax = (tAmax < tBmax) ? tAmax : tBmax;

    float length = tmax - tmin;
    if (length <= 0.0) {
        if (Px0) *Px0 = 0.0; if (Py0) *Py0 = 0.0; if (Pz0) *Pz0 = 0.0;
        if (Px1) *Px1 = 0.0; if (Py1) *Py1 = 0.0; if (Pz1) *Pz1 = 0.0;
        return -fabsf(length); // no overlap
    }

    // Segment endpoints
    if (Px0) *Px0 = Px + tmin * Qx;
    if (Py0) *Py0 = Py + tmin * Qy;
    if (Pz0) *Pz0 = Pz + tmin * Qz;

    if (Px1) *Px1 = Px + tmax * Qx;
    if (Py1) *Py1 = Py + tmax * Qy;
    if (Pz1) *Pz1 = Pz + tmax * Qz;

    return length;
}



extern float TriTriSegmentFastLen(
    // Triangle A
    float Ax1, float Ay1, float Az1,
    float Ax2, float Ay2, float Az2,
    float Ax3, float Ay3, float Az3,
    // Triangle B
    float Bx1, float By1, float Bz1,
    float Bx2, float By2, float Bz2,
    float Bx3, float By3, float Bz3)
{
    // Normals A and B
    float nAx = (Ay2 - Ay1) * (Az3 - Az1) - (Az2 - Az1) * (Ay3 - Ay1);
    float nAy = (Az2 - Az1) * (Ax3 - Ax1) - (Ax2 - Ax1) * (Az3 - Az1);
    float nAz = (Ax2 - Ax1) * (Ay3 - Ay1) - (Ay2 - Ay1) * (Ax3 - Ax1);

    float nBx = (By2 - By1) * (Bz3 - Bz1) - (Bz2 - Bz1) * (By3 - By1);
    float nBy = (Bz2 - Bz1) * (Bx3 - Bx1) - (Bx2 - Bx1) * (Bz3 - Bz1);
    float nBz = (Bx2 - Bx1) * (By3 - By1) - (By2 - By1) * (Bx3 - Bx1);

    // Plane constants
    float dA = nAx * Ax1 + nAy * Ay1 + nAz * Az1;
    float dB = nBx * Bx1 + nBy * By1 + nBz * Bz1;

    // Line direction Q = nA × nB
    float Qx = nAy * nBz - nAz * nBy;
    float Qy = nAz * nBx - nAx * nBz;
    float Qz = nAx * nBy - nAy * nBx;

    // Point P = ((dA nB - dB nA) × Q) / |Q|²
    float Px = (dA * nBx - dB * nAx);
    float Py = (dA * nBy - dB * nAy);
    float Pz = (dA * nBz - dB * nAz);

    float tmpx = Py * Qz - Pz * Qy;
    float tmpy = Pz * Qx - Px * Qz;
    float tmpz = Px * Qy - Py * Qx;

    float denom = Qx * Qx + Qy * Qy + Qz * Qz;
    if (denom == 0.0) return 0.0; // parallel planes

    Px = tmpx / denom;
    Py = tmpy / denom;
    Pz = tmpz / denom;

    // Project vertices of A
    float tAmin = ((Ax1 - Px) * Qx + (Ay1 - Py) * Qy + (Az1 - Pz) * Qz) / denom;
    float tAmax = tAmin;

    float t = ((Ax2 - Px) * Qx + (Ay2 - Py) * Qy + (Az2 - Pz) * Qz) / denom;
    if (t < tAmin) tAmin = t; else if (t > tAmax) tAmax = t;

    t = ((Ax3 - Px) * Qx + (Ay3 - Py) * Qy + (Az3 - Pz) * Qz) / denom;
    if (t < tAmin) tAmin = t; else if (t > tAmax) tAmax = t;

    // Project vertices of B
    float tBmin = ((Bx1 - Px) * Qx + (By1 - Py) * Qy + (Bz1 - Pz) * Qz) / denom;
    float tBmax = tBmin;

    t = ((Bx2 - Px) * Qx + (By2 - Py) * Qy + (Bz2 - Pz) * Qz) / denom;
    if (t < tBmin) tBmin = t; else if (t > tBmax) tBmax = t;

    t = ((Bx3 - Px) * Qx + (By3 - Py) * Qy + (Bz3 - Pz) * Qz) / denom;
    if (t < tBmin) tBmin = t; else if (t > tBmax) tBmax = t;

    // Overlap interval
    float tmin = (tAmin > tBmin) ? tAmin : tBmin;
    float tmax = (tAmax < tBmax) ? tAmax : tBmax;

    float length = tmax - tmin;
    if (length <= 0.0) return -fabsf(length); // no overlap
    return length;
}