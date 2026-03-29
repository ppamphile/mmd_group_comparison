"""
Calcul de la MMD avec noyau additif.
"""

import numpy as np


def laplace_kernel(X, Y, bandwidth):
    dists = np.abs(X[:, None, :] - Y[None, :, :]).sum(axis=2)
    return np.exp(-dists / bandwidth)


def compute_block_kernel(X, Y, block, bandwidth):

    Xb = X[block.columns].values
    Yb = Y[block.columns].values

    return laplace_kernel(Xb, Yb, bandwidth)


def estimate_block_bandwidths(df, config):
    """
    Estimation simple des bandwidths (médiane des distances).
    """
    bandwidths = {}

    for block in config.blocks:
        data = df[block.columns].values

        # distance L1
        dists = np.abs(data[:, None, :] - data[None, :, :]).sum(axis=2)

        bandwidth = np.median(dists[dists > 0])
        bandwidths[block.name] = bandwidth if bandwidth > 0 else 1.0

    return bandwidths


def compute_mmd_block(Kxx, Kyy, Kxy):
    m = Kxx.shape[0]
    n = Kyy.shape[0]

    return (
        (Kxx.sum() - np.trace(Kxx)) / (m * (m - 1))
        + (Kyy.sum() - np.trace(Kyy)) / (n * (n - 1))
        - 2 * Kxy.mean()
    )


def compute_mmd_decomposition(X, Y, config, bandwidths):

    total = 0.0
    contribs = {}

    for block in config.blocks:
        bw = bandwidths[block.name]

        Kxx = compute_block_kernel(X, X, block, bw)
        Kyy = compute_block_kernel(Y, Y, block, bw)
        Kxy = compute_block_kernel(X, Y, block, bw)

        mmd_block = compute_mmd_block(Kxx, Kyy, Kxy)

        contribs[block.name] = mmd_block
        total += mmd_block * block.weight

    return total, contribs