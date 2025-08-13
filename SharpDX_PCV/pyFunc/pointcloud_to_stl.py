#!/usr/bin/env python3
"""
点云到STL转换工具
使用Open3D库进行高质量的三角剖分和STL生成
"""

import sys
import os
import json
import numpy as np
import open3d as o3d
from pathlib import Path
import argparse
import time
from scipy.spatial import cKDTree
from typing import List, Tuple, Optional, Dict, Any


class PointCloudToSTLConverter:
    """点云到STL转换器 - 支持平面检测和四角化优化"""

    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback

        # 平面检测参数（针对毫米单位数据优化）
        self.plane_config = {
            'distance_threshold': 0.1,      # 平面距离阈值 (mm)
            'min_plane_points': 500,        # 最小平面点数
            'ransac_n': 3,                  # RANSAC采样点数
            'num_iterations': 5000,         # RANSAC迭代次数
            'max_planes': 10,               # 最大平面数量
        }

        # 重建参数
        self.recon_config = {
            'method': 'hybrid',             # 混合方法：先平面四角化，再重建残余
            'merge_epsilon': 1e-3,          # 顶点合并阈值 (mm)
        }
        
    def log_progress(self, progress, message):
        """记录进度"""
        if self.progress_callback:
            self.progress_callback(progress, message)
        else:
            print(f"[{progress:3.0f}%] {message}")
    
    def load_txt_files(self, file_paths):
        """加载多个TXT文件"""
        self.log_progress(5, "开始加载TXT文件...")
        
        all_points = []
        
        for i, file_path in enumerate(file_paths):
            self.log_progress(5 + (i * 15 / len(file_paths)), 
                            f"加载文件 {i+1}/{len(file_paths)}: {Path(file_path).name}")
            
            try:
                # 尝试不同的分隔符
                for delimiter in [' ', '\t', ',']:
                    try:
                        data = np.loadtxt(file_path, delimiter=delimiter, comments=['#', '//'])
                        if data.size > 0:
                            break
                    except:
                        continue
                else:
                    # 如果所有分隔符都失败，尝试混合分隔符
                    points = []
                    with open(file_path, 'r', encoding='utf-8') as f:
                        for line in f:
                            line = line.strip()
                            if not line or line.startswith('#') or line.startswith('//'):
                                continue
                            
                            # 分割行，支持多种分隔符
                            parts = line.replace(',', ' ').replace('\t', ' ').split()
                            if len(parts) >= 2:
                                try:
                                    x = float(parts[0])
                                    y = float(parts[1])
                                    z = float(parts[2]) if len(parts) >= 3 else 0.0
                                    points.append([x, y, z])
                                except ValueError:
                                    continue
                    
                    if points:
                        data = np.array(points)
                    else:
                        print(f"警告: 无法解析文件 {file_path}")
                        continue
                
                # 确保数据是3D的
                if data.ndim == 1:
                    data = data.reshape(1, -1)
                
                if data.shape[1] < 3:
                    # 如果只有2列，添加Z=0
                    if data.shape[1] == 2:
                        z_column = np.zeros((data.shape[0], 1))
                        data = np.hstack([data, z_column])
                    else:
                        print(f"警告: 文件 {file_path} 数据格式不正确")
                        continue
                elif data.shape[1] > 3:
                    # 如果超过3列，只取前3列
                    data = data[:, :3]
                
                all_points.append(data)
                
            except Exception as e:
                print(f"错误: 无法加载文件 {file_path}: {e}")
                continue
        
        if not all_points:
            raise ValueError("没有成功加载任何点云数据")
        
        # 合并所有点
        combined_points = np.vstack(all_points)
        self.log_progress(20, f"成功加载 {len(combined_points)} 个点")
        
        return combined_points
    
    def create_point_cloud(self, points):
        """创建Open3D点云对象"""
        self.log_progress(25, "创建点云对象...")
        
        # 创建Open3D点云
        pcd = o3d.geometry.PointCloud()
        pcd.points = o3d.utility.Vector3dVector(points)
        
        return pcd
    
    def _estimate_spacing(self, points: np.ndarray, sample_size: int = 5000) -> float:
        """估计典型点间距（鲁棒版，优先使用SciPy cKDTree）"""
        try:
            if points is None or len(points) < 2:
                return 1e-2
            pts = np.asarray(points, dtype=float)
            # 去除非有限值
            mask = np.isfinite(pts).all(axis=1)
            pts = pts[mask]
            if len(pts) < 2:
                return 1e-2

            m = min(sample_size, len(pts))
            idx = np.random.choice(len(pts), size=m, replace=False)
            sample = pts[idx]

            # 使用 cKDTree 计算最近邻距离（k=2：自身+最近邻）
            tree = cKDTree(pts, leafsize=64)
            try:
                dists, _ = tree.query(sample, k=2, workers=-1)
            except TypeError:
                # 某些 SciPy 版本没有 workers 参数
                dists, _ = tree.query(sample, k=2)

            nn = dists[:, 1]  # 最近邻距离
            nn = nn[np.isfinite(nn) & (nn > 0)]
            if nn.size == 0:
                return 1e-2
            return float(np.median(nn))
        except Exception:
            # 退化：用包围盒估计（保守）
            try:
                bb = np.ptp(points, axis=0)
                approx = (np.cbrt(np.prod(bb) / max(len(points), 1))) if np.all(bb > 0) else max(bb.max() / 100.0, 1e-2)
                return float(max(approx, 1e-3))
            except Exception:
                return 1e-2

    def preprocess_point_cloud(self, pcd):
        """预处理点云（自适应密度）"""
        self.log_progress(30, "预处理点云...")

        # 估计点间距并做可选下采样（防止过密）
        pts = np.asarray(pcd.points)
        spacing = max(self._estimate_spacing(pts), 1e-4)
        voxel = max(spacing * 0.8, 1e-4)
        # 适度控制点数规模（>30万时也做降采样）
        if len(pts) > 300_000:
            self.log_progress(32, f"点数较多，进行体素下采样 voxel={voxel:.5f}")
            pcd = pcd.voxel_down_sample(voxel)
            pts = np.asarray(pcd.points)  # 更新缓存

        # 移除重复点
        pcd = pcd.remove_duplicated_points()
        self.log_progress(35, f"移除重复点后剩余 {len(pcd.points)} 个点")

        # 自适应离群点移除强度
        nb_neighbors = 30 if len(pcd.points) > 200_000 else 20
        std_ratio = 2.0 if len(pcd.points) > 200_000 else 1.5
        pcd, _ = pcd.remove_statistical_outlier(nb_neighbors=nb_neighbors, std_ratio=std_ratio)
        self.log_progress(40, f"移除离群点后剩余 {len(pcd.points)} 个点")

        # 自适应法向量估计半径
        normal_radius = max(spacing * 3.0, 1e-3)
        max_nn = 80 if len(pcd.points) > 500_000 else (60 if len(pcd.points) > 200_000 else 30)
        try:
            pcd.estimate_normals(search_param=o3d.geometry.KDTreeSearchParamHybrid(radius=normal_radius, max_nn=max_nn))
        except RuntimeError:
            # 退化：使用 KNN 模式
            pcd.estimate_normals(search_param=o3d.geometry.KDTreeSearchParamKNN(knn=min(30, len(pcd.points))))
        self.log_progress(45, f"法向量估计完成 (r={normal_radius:.5f}, nn={max_nn})")

        # 保持法向量一致方向（防止失败）
        try:
            pcd.orient_normals_consistent_tangent_plane(k=min(30, len(pcd.points)))
        except Exception:
            pass

        return pcd

    def detect_planes(self, pcd):
        """使用RANSAC检测平面"""
        self.log_progress(47, "开始平面检测...")

        planes = []
        residual_pcd = pcd

        for i in range(self.plane_config['max_planes']):
            if len(residual_pcd.points) < self.plane_config['min_plane_points']:
                break

            try:
                # RANSAC平面分割
                plane_model, inliers = residual_pcd.segment_plane(
                    distance_threshold=self.plane_config['distance_threshold'],
                    ransac_n=self.plane_config['ransac_n'],
                    num_iterations=self.plane_config['num_iterations']
                )

                if len(inliers) < self.plane_config['min_plane_points']:
                    break

                # 提取平面点云
                plane_pcd = residual_pcd.select_by_index(inliers)
                planes.append({
                    'model': plane_model,  # [a, b, c, d] 平面方程 ax+by+cz+d=0
                    'pcd': plane_pcd,
                    'inliers_count': len(inliers)
                })

                # 移除已检测的平面点
                residual_pcd = residual_pcd.select_by_index(inliers, invert=True)

                self.log_progress(47 + i, f"检测到平面 {i+1}: {len(inliers)} 个点")

            except Exception as e:
                self.log_progress(47 + i, f"平面检测失败: {e}")
                break

        self.log_progress(50, f"平面检测完成，共检测到 {len(planes)} 个平面，剩余 {len(residual_pcd.points)} 个点")
        return planes, residual_pcd

    def extract_plane_quad_corners(self, plane_pcd, plane_model):
        """提取平面的四个角点"""
        try:
            # 获取平面法向量
            a, b, c, d = plane_model
            normal = np.array([a, b, c])
            normal = normal / np.linalg.norm(normal)

            # 获取平面中心点
            points = np.asarray(plane_pcd.points)
            center = np.mean(points, axis=0)

            # 构建局部坐标系
            # 选择一个与法向量不平行的向量
            if abs(normal[2]) < 0.9:
                temp = np.array([0, 0, 1])
            else:
                temp = np.array([1, 0, 0])

            # 构建正交基
            u = np.cross(normal, temp)
            u = u / np.linalg.norm(u)
            v = np.cross(normal, u)

            # 将3D点投影到2D平面
            points_2d = []
            for point in points:
                relative = point - center
                x_2d = np.dot(relative, u)
                y_2d = np.dot(relative, v)
                points_2d.append([x_2d, y_2d])

            points_2d = np.array(points_2d)

            # 计算2D边界框（简化版最小外接矩形）
            min_x, max_x = np.min(points_2d[:, 0]), np.max(points_2d[:, 0])
            min_y, max_y = np.min(points_2d[:, 1]), np.max(points_2d[:, 1])

            # 四个角点（2D）
            corners_2d = np.array([
                [min_x, min_y],
                [max_x, min_y],
                [max_x, max_y],
                [min_x, max_y]
            ])

            # 投影回3D
            corners_3d = []
            for corner_2d in corners_2d:
                corner_3d = center + corner_2d[0] * u + corner_2d[1] * v
                corners_3d.append(corner_3d)

            return np.array(corners_3d)

        except Exception as e:
            self.log_progress(48, f"四角点提取失败: {e}")
            # 退化方案：使用包围盒
            bbox = plane_pcd.get_axis_aligned_bounding_box()
            return np.asarray(bbox.get_box_points())[:4]  # 取前4个点

    def create_quad_mesh(self, corners):
        """从四个角点创建两个三角形"""
        try:
            # 创建网格
            mesh = o3d.geometry.TriangleMesh()
            mesh.vertices = o3d.utility.Vector3dVector(corners)

            # 创建两个三角形（沿短对角线分割）
            # 计算对角线长度
            diag1_len = np.linalg.norm(corners[2] - corners[0])
            diag2_len = np.linalg.norm(corners[3] - corners[1])

            if diag1_len < diag2_len:
                # 沿对角线 0-2 分割
                triangles = [[0, 1, 2], [0, 2, 3]]
            else:
                # 沿对角线 1-3 分割
                triangles = [[0, 1, 3], [1, 2, 3]]

            mesh.triangles = o3d.utility.Vector3iVector(triangles)
            mesh.compute_vertex_normals()

            return mesh

        except Exception as e:
            self.log_progress(48, f"四角网格创建失败: {e}")
            return None

    def create_mesh(self, pcd):
        """创建三角网格（混合方法：平面四角化 + 残余重建）"""
        self.log_progress(46, "开始混合网格重建...")

        # 1. 平面检测
        planes, residual_pcd = self.detect_planes(pcd)

        all_meshes = []

        # 2. 为每个平面创建四角网格
        for i, plane in enumerate(planes):
            self.log_progress(50 + i, f"处理平面 {i+1}/{len(planes)}")
            try:
                corners = self.extract_plane_quad_corners(plane['pcd'], plane['model'])
                quad_mesh = self.create_quad_mesh(corners)
                if quad_mesh is not None:
                    all_meshes.append(quad_mesh)
                    self.log_progress(50 + i, f"平面 {i+1} 四角化完成: 2个三角形")
            except Exception as e:
                self.log_progress(50 + i, f"平面 {i+1} 四角化失败: {e}")

        # 3. 对残余点进行传统重建
        residual_mesh = None
        if len(residual_pcd.points) > 100:  # 只有足够的点才进行重建
            self.log_progress(60, f"重建残余点云: {len(residual_pcd.points)} 个点")
            residual_mesh = self._reconstruct_residual_points(residual_pcd)
            if residual_mesh is not None:
                all_meshes.append(residual_mesh)

        # 4. 合并所有网格
        if not all_meshes:
            raise ValueError("无法生成任何有效网格")

        self.log_progress(70, f"合并 {len(all_meshes)} 个网格...")
        final_mesh = self._merge_meshes(all_meshes)

        self.log_progress(72, f"混合重建完成: {len(final_mesh.triangles)} 个三角形")
        return final_mesh

    def _reconstruct_residual_points(self, pcd):
        """重建残余点云（非平面部分）"""
        if len(pcd.points) < 50:
            return None

        pts = np.asarray(pcd.points)
        spacing = self._estimate_spacing(pts)

        # 优先尝试 Alpha-Shape（对残余点更适合）
        try:
            alpha = max(spacing * 2.0, 5e-4)
            mesh = o3d.geometry.TriangleMesh.create_from_point_cloud_alpha_shape(pcd, alpha=alpha)
            if len(mesh.triangles) > 0:
                self.log_progress(65, f"残余点Alpha-Shape成功: {len(mesh.triangles)} 个三角形")
                return mesh
        except Exception as e:
            self.log_progress(62, f"Alpha-Shape失败: {e}")

        # 备用：Ball Pivoting
        try:
            base = max(spacing, 1e-3)
            radii = [base * r for r in (0.8, 1.2, 2.0)]
            mesh = o3d.geometry.TriangleMesh.create_from_point_cloud_ball_pivoting(
                pcd, o3d.utility.DoubleVector(radii)
            )
            if len(mesh.triangles) > 0:
                self.log_progress(65, f"残余点Ball Pivoting成功: {len(mesh.triangles)} 个三角形")
                return mesh
        except Exception as e:
            self.log_progress(63, f"Ball Pivoting失败: {e}")

        # 最后尝试 Poisson（可能过度平滑）
        try:
            depth = 8  # 较小的深度避免过度平滑
            mesh, _ = o3d.geometry.TriangleMesh.create_from_point_cloud_poisson(
                pcd, depth=depth, width=0, scale=1.1, linear_fit=False
            )
            if len(mesh.triangles) > 0:
                # 裁剪到点云范围
                aabb = pcd.get_axis_aligned_bounding_box()
                mesh = mesh.crop(aabb.scale(1.1, aabb.get_center()))
                self.log_progress(65, f"残余点Poisson成功: {len(mesh.triangles)} 个三角形")
                return mesh
        except Exception as e:
            self.log_progress(64, f"Poisson失败: {e}")

        return None

    def _merge_meshes(self, meshes):
        """合并多个网格"""
        if len(meshes) == 1:
            return meshes[0]

        # 合并顶点和三角形
        all_vertices = []
        all_triangles = []
        vertex_offset = 0

        for mesh in meshes:
            vertices = np.asarray(mesh.vertices)
            triangles = np.asarray(mesh.triangles)

            all_vertices.append(vertices)
            # 调整三角形索引
            adjusted_triangles = triangles + vertex_offset
            all_triangles.append(adjusted_triangles)
            vertex_offset += len(vertices)

        # 创建合并后的网格
        merged_mesh = o3d.geometry.TriangleMesh()
        merged_mesh.vertices = o3d.utility.Vector3dVector(np.vstack(all_vertices))
        merged_mesh.triangles = o3d.utility.Vector3iVector(np.vstack(all_triangles))

        # 合并重复顶点
        merged_mesh.merge_close_vertices(self.recon_config['merge_epsilon'])

        return merged_mesh
    
    def postprocess_mesh(self, mesh):
        """后处理网格"""
        self.log_progress(75, "后处理网格...")

        # 移除重复的三角形
        mesh.remove_duplicated_triangles()

        # 移除退化的三角形
        mesh.remove_degenerate_triangles()

        # 移除非流形边
        mesh.remove_non_manifold_edges()

        # 计算顶点法向量（STL保存需要）
        mesh.compute_vertex_normals()

        # 可选：裁剪孤立组件（去除浮动碎片）
        try:
            cc = mesh.cluster_connected_triangles()[0]
            if len(cc) > 0:
                counts = np.bincount(cc)
                keep = counts.argmax()
                mask = (cc == keep)
                triangles = np.asarray(mesh.triangles)
                mesh.triangles = o3d.utility.Vector3iVector(triangles[mask])
                mesh.remove_unreferenced_vertices()
        except Exception:
            pass

        self.log_progress(80, f"后处理完成，最终 {len(mesh.triangles)} 个三角形")

        return mesh
    
    def save_stl(self, mesh, output_path):
        """保存STL文件（确保法向量计算）"""
        self.log_progress(90, f"保存STL文件: {output_path}")

        # 确保输出目录存在
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # 确保法向量存在（Open3D保存STL前必需）
        try:
            if not mesh.has_vertex_normals():
                mesh.compute_vertex_normals()
        except Exception:
            mesh.compute_vertex_normals()

        # 作为额外保险，计算三角形法向量
        if not mesh.has_triangle_normals():
            try:
                mesh.compute_triangle_normals()
            except Exception:
                pass

        # 保存为STL格式
        success = o3d.io.write_triangle_mesh(str(output_path), mesh)

        if not success:
            raise ValueError(f"无法保存STL文件到 {output_path}")

        self.log_progress(100, f"STL文件保存成功: {output_path}")

        return str(output_path)
    
    def convert(self, input_files, output_dir, output_name=None):
        """执行完整的转换流程"""
        try:
            # 1. 加载点云数据
            points = self.load_txt_files(input_files)

            # 2. 创建点云对象
            pcd = self.create_point_cloud(points)

            # 3. 预处理
            pcd = self.preprocess_point_cloud(pcd)

            # 4. 创建网格
            mesh = self.create_mesh(pcd)

            # 5. 后处理
            mesh = self.postprocess_mesh(mesh)

            # 6. 构建输出路径
            out_dir = Path(output_dir) if output_dir else Path.cwd()
            out_dir.mkdir(parents=True, exist_ok=True)
            if output_name is None:
                ts = time.strftime('%Y%m%d_%H%M%S')
                output_name = f"pointcloud_{ts}.stl"
            out_path = str(out_dir / output_name)

            # 7. 保存STL
            result_path = self.save_stl(mesh, out_path)
            
            return {
                'success': True,
                'output_path': result_path,
                'original_points': len(points),
                'final_triangles': len(mesh.triangles),
                'final_vertices': len(mesh.vertices)
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description='点云到STL转换工具')
    parser.add_argument('input_files', nargs='+', help='输入的TXT文件路径')
    parser.add_argument('-o', '--output-dir', required=False, help='输出目录路径')
    parser.add_argument('--output-name', required=False, help='输出文件名（可选，默认自动命名）')
    parser.add_argument('--json-output', help='JSON格式的结果输出文件')

    args = parser.parse_args()

    # 创建转换器
    converter = PointCloudToSTLConverter()

    # 执行转换
    result = converter.convert(args.input_files, args.output_dir, args.output_name)
    
    # 输出结果
    if args.json_output:
        with open(args.json_output, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=2, ensure_ascii=False)
    
    if result['success']:
        print(f"转换成功!")
        print(f"输出文件: {result['output_path']}")
        print(f"原始点数: {result['original_points']:,}")
        print(f"最终三角形数: {result['final_triangles']:,}")
        print(f"最终顶点数: {result['final_vertices']:,}")
        sys.exit(0)
    else:
        print(f"转换失败: {result['error']}")
        sys.exit(1)


if __name__ == '__main__':
    main()
